VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMctl32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmTcpServer 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  '���� ����
   Caption         =   "ParkingManager��"
   ClientHeight    =   11250
   ClientLeft      =   7080
   ClientTop       =   2895
   ClientWidth     =   19320
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
   ScaleHeight     =   11250
   ScaleWidth      =   19320
   Begin VB.CommandButton Command9 
      Caption         =   "CCTV"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4740
      TabIndex        =   361
      Top             =   0
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Ž����"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   22950
      TabIndex        =   359
      Top             =   1425
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.CommandButton Command5 
      Caption         =   "��������"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   22950
      TabIndex        =   358
      Top             =   975
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.CommandButton Command7 
      Caption         =   "���â"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   22950
      TabIndex        =   357
      Top             =   1875
      Visible         =   0   'False
      Width           =   945
   End
   Begin MSWinsockLib.Winsock WinsockS_Devices 
      Left            =   13455
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin LPR_PARKING_HOST.Server Server_WebDC 
      Left            =   5715
      Top             =   -15
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin MSWinsockLib.Winsock Winsock_GateAgentR 
      Index           =   0
      Left            =   20310
      Top             =   3585
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin LPR_PARKING_HOST.Server Server_GateAgentR 
      Index           =   0
      Left            =   20310
      Top             =   3150
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.CommandButton Command4 
      Caption         =   "�����Ʈ"
      Enabled         =   0   'False
      Height          =   435
      Left            =   22950
      TabIndex        =   340
      Top             =   525
      Visible         =   0   'False
      Width           =   945
   End
   Begin MSWinsockLib.Winsock WinsockS_CertPC 
      Left            =   13935
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer Timer_Emerg_Vertical 
      Enabled         =   0   'False
      Index           =   5
      Interval        =   100
      Left            =   12795
      Top             =   0
   End
   Begin VB.Timer Timer_Emerg_Vertical 
      Enabled         =   0   'False
      Index           =   4
      Interval        =   100
      Left            =   12375
      Top             =   0
   End
   Begin VB.Timer Timer_Emerg_Vertical 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   100
      Left            =   11955
      Top             =   0
   End
   Begin VB.Timer Timer_Emerg_Vertical 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   100
      Left            =   11535
      Top             =   0
   End
   Begin VB.Timer Timer_Emerg_Vertical 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   100
      Left            =   11115
      Top             =   0
   End
   Begin VB.Timer Timer_Emerg_Vertical 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   100
      Left            =   10695
      Top             =   0
   End
   Begin VB.Timer Timer_ParkFullLight 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   17880
      Top             =   0
   End
   Begin MSWinsockLib.Winsock ParkFullLightS_sock 
      Left            =   17460
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ComboBox cmb_DeviceMode 
      Height          =   330
      Index           =   5
      ItemData        =   "FrmTcpServer.frx":0000
      Left            =   10545
      List            =   "FrmTcpServer.frx":000A
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   323
      Top             =   13050
      Visible         =   0   'False
      Width           =   1600
   End
   Begin VB.ComboBox cmb_DeviceMode 
      Height          =   330
      Index           =   4
      ItemData        =   "FrmTcpServer.frx":0018
      Left            =   7815
      List            =   "FrmTcpServer.frx":0022
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   322
      Top             =   13110
      Visible         =   0   'False
      Width           =   1600
   End
   Begin VB.ComboBox cmb_DeviceMode 
      Height          =   330
      Index           =   3
      ItemData        =   "FrmTcpServer.frx":0030
      Left            =   5355
      List            =   "FrmTcpServer.frx":003A
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   321
      Top             =   13095
      Visible         =   0   'False
      Width           =   1600
   End
   Begin VB.ComboBox cmb_DeviceMode 
      Height          =   330
      Index           =   2
      ItemData        =   "FrmTcpServer.frx":0048
      Left            =   2685
      List            =   "FrmTcpServer.frx":0052
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   320
      Top             =   13230
      Visible         =   0   'False
      Width           =   1600
   End
   Begin VB.ComboBox cmb_DeviceMode 
      Height          =   330
      Index           =   1
      ItemData        =   "FrmTcpServer.frx":0060
      Left            =   60
      List            =   "FrmTcpServer.frx":006A
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   319
      Top             =   12645
      Visible         =   0   'False
      Width           =   1600
   End
   Begin VB.ComboBox cmb_DisplayMode 
      Height          =   330
      Index           =   5
      ItemData        =   "FrmTcpServer.frx":0078
      Left            =   13020
      List            =   "FrmTcpServer.frx":0082
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   318
      Top             =   12645
      Visible         =   0   'False
      Width           =   1600
   End
   Begin VB.ComboBox cmb_DisplayMode 
      Height          =   330
      Index           =   4
      ItemData        =   "FrmTcpServer.frx":0090
      Left            =   10545
      List            =   "FrmTcpServer.frx":009A
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   317
      Top             =   12705
      Visible         =   0   'False
      Width           =   1600
   End
   Begin VB.ComboBox cmb_DisplayMode 
      Height          =   330
      Index           =   3
      ItemData        =   "FrmTcpServer.frx":00A8
      Left            =   7830
      List            =   "FrmTcpServer.frx":00B2
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   316
      Top             =   12750
      Visible         =   0   'False
      Width           =   1600
   End
   Begin VB.ComboBox cmb_DisplayMode 
      Height          =   330
      Index           =   2
      ItemData        =   "FrmTcpServer.frx":00C0
      Left            =   5370
      List            =   "FrmTcpServer.frx":00CA
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   315
      Top             =   12750
      Visible         =   0   'False
      Width           =   1600
   End
   Begin VB.ComboBox cmb_DisplayMode 
      Height          =   330
      Index           =   1
      ItemData        =   "FrmTcpServer.frx":00D8
      Left            =   2700
      List            =   "FrmTcpServer.frx":00E2
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   314
      Top             =   12870
      Visible         =   0   'False
      Width           =   1600
   End
   Begin MSWinsockLib.Winsock DBSock 
      Left            =   18855
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer GateTimer 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   200
      Left            =   7575
      Top             =   0
   End
   Begin VB.Timer GateTimer 
      Enabled         =   0   'False
      Index           =   1
      Interval        =   200
      Left            =   7995
      Top             =   0
   End
   Begin VB.Timer GateTimer 
      Enabled         =   0   'False
      Index           =   2
      Interval        =   200
      Left            =   8430
      Top             =   0
   End
   Begin VB.Timer GateTimer 
      Enabled         =   0   'False
      Index           =   3
      Interval        =   200
      Left            =   8865
      Top             =   0
   End
   Begin VB.Timer GateTimer 
      Enabled         =   0   'False
      Index           =   4
      Interval        =   200
      Left            =   9285
      Top             =   0
   End
   Begin VB.Timer GateTimer 
      Enabled         =   0   'False
      Index           =   5
      Interval        =   200
      Left            =   9705
      Top             =   0
   End
   Begin VB.Timer DBTimer 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   18435
      Top             =   0
   End
   Begin MSWinsockLib.Winsock FreepassR_sock 
      Left            =   16890
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock FreepassS_sock 
      Left            =   16455
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock MobileR_Sock 
      Left            =   15840
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin LPR_PARKING_HOST.Server Server 
      Left            =   6150
      Top             =   -15
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin MSWinsockLib.Winsock Aps_UDP 
      Left            =   15210
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock ApsS_sock 
      Left            =   14775
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CheckBox chk_ApsYN 
      BackColor       =   &H00E0E0E0&
      Caption         =   "���"
      Height          =   225
      Left            =   18675
      TabIndex        =   26
      ToolTipText     =   "���üũ: ���ܱ� ���� ������ ��������Ⱑ ���Ե˴ϴ�."
      Top             =   1500
      Width           =   630
   End
   Begin VB.CheckBox chk_PreApsYN 
      BackColor       =   &H00E0E0E0&
      Caption         =   "���"
      Height          =   225
      Left            =   16785
      TabIndex        =   2
      ToolTipText     =   "���üũ: ���ܱ� ���� ������ ��������Ⱑ ���Ե˴ϴ�."
      Top             =   1500
      Width           =   660
   End
   Begin VB.CheckBox chk_RemoteYN 
      BackColor       =   &H00E0E0E0&
      Caption         =   "���"
      Height          =   225
      Index           =   1
      Left            =   14505
      TabIndex        =   108
      ToolTipText     =   "�PC �̸鼭, ����͸�PC�� �����͸� �����ؾ� �Ѵٸ� ""���"" üũ �ϼ���"
      Top             =   1500
      Width           =   690
   End
   Begin VB.CheckBox chk_RemoteYN 
      BackColor       =   &H00E0E0E0&
      Caption         =   "���"
      Height          =   210
      Index           =   0
      Left            =   12795
      TabIndex        =   107
      ToolTipText     =   "����͸�PC �뵵�� ����Ѵٸ�  ""���"" üũ �ϼ���"
      Top             =   1500
      Width           =   690
   End
   Begin VB.CheckBox chk_UseYN 
      BackColor       =   &H00E0E0E0&
      Caption         =   "���"
      Height          =   210
      Index           =   5
      Left            =   18600
      TabIndex        =   285
      Top             =   2775
      Width           =   630
   End
   Begin VB.CheckBox chk_UseYN 
      BackColor       =   &H00E0E0E0&
      Caption         =   "���"
      Height          =   210
      Index           =   4
      Left            =   15390
      TabIndex        =   256
      Top             =   2775
      Width           =   630
   End
   Begin VB.CheckBox chk_UseYN 
      BackColor       =   &H00E0E0E0&
      Caption         =   "���"
      Height          =   210
      Index           =   3
      Left            =   12180
      TabIndex        =   227
      Top             =   2775
      Width           =   630
   End
   Begin VB.CheckBox chk_UseYN 
      BackColor       =   &H00E0E0E0&
      Caption         =   "���"
      Height          =   210
      Index           =   2
      Left            =   8970
      TabIndex        =   198
      Top             =   2775
      Width           =   630
   End
   Begin VB.CheckBox chk_UseYN 
      BackColor       =   &H00E0E0E0&
      Caption         =   "���"
      Height          =   210
      Index           =   1
      Left            =   5760
      TabIndex        =   169
      Top             =   2775
      Width           =   630
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   " LANE6 "
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   5100
      Index           =   5
      Left            =   16110
      TabIndex        =   286
      Top             =   2910
      Width           =   3195
      Begin VB.CommandButton cmd_GateTestDown 
         Caption         =   "����"
         Height          =   330
         Index           =   5
         Left            =   885
         TabIndex        =   368
         Top             =   4680
         Width           =   690
      End
      Begin VB.ComboBox cmb_DispShiftSpeed 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   5
         ItemData        =   "FrmTcpServer.frx":00F0
         Left            =   1725
         List            =   "FrmTcpServer.frx":010F
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   352
         ToolTipText     =   "������ �̵� �ӵ�����(���ڰ� ���� ���� ���� �̵��մϴ�), ���� �� �Ϲݹ�ư ��������"
         Top             =   4695
         Width           =   615
      End
      Begin VB.CommandButton cmd_NmlShift 
         Caption         =   "�̵�"
         Height          =   330
         Index           =   5
         Left            =   2415
         TabIndex        =   336
         ToolTipText     =   "���ι��� �Ϲ������� 6���ڱ��� �����մϴ�"
         Top             =   4680
         Width           =   690
      End
      Begin VB.CheckBox chk_BackCam_YN 
         Caption         =   "Check3"
         Height          =   225
         Index           =   5
         Left            =   885
         TabIndex        =   303
         Top             =   1350
         Width           =   195
      End
      Begin VB.ComboBox cmb_PrintModel 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         ItemData        =   "FrmTcpServer.frx":0137
         Left            =   1695
         List            =   "FrmTcpServer.frx":0141
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   302
         Top             =   1320
         Width           =   945
      End
      Begin VB.ComboBox cmb_PrintPort 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         ItemData        =   "FrmTcpServer.frx":0155
         Left            =   1695
         List            =   "FrmTcpServer.frx":019E
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   301
         Top             =   1680
         Width           =   945
      End
      Begin VB.CheckBox chk_GuestYN 
         Caption         =   "Check3"
         Height          =   225
         Index           =   5
         Left            =   885
         TabIndex        =   300
         Top             =   1740
         Width           =   195
      End
      Begin VB.ComboBox CmbScreen 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         ItemData        =   "FrmTcpServer.frx":0237
         Left            =   885
         List            =   "FrmTcpServer.frx":0239
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   299
         Top             =   930
         Width           =   1755
      End
      Begin VB.ComboBox cmb_Inout 
         Height          =   330
         Index           =   5
         ItemData        =   "FrmTcpServer.frx":023B
         Left            =   885
         List            =   "FrmTcpServer.frx":0245
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   298
         Top             =   240
         Width           =   1725
      End
      Begin VB.CommandButton cmd_GateTest 
         Caption         =   "���ܱ�"
         Height          =   330
         Index           =   5
         Left            =   885
         TabIndex        =   297
         Top             =   4290
         Width           =   690
      End
      Begin VB.CommandButton cmd_EmgTest 
         Caption         =   "���"
         Height          =   330
         Index           =   5
         Left            =   1650
         TabIndex        =   296
         Top             =   4290
         Width           =   690
      End
      Begin VB.CommandButton cmd_NmlTest 
         Caption         =   "����"
         Height          =   330
         Index           =   5
         Left            =   2415
         TabIndex        =   295
         Top             =   4290
         Width           =   690
      End
      Begin VB.CommandButton cmd_CapTest 
         Caption         =   "ĸ��"
         Height          =   330
         Index           =   5
         Left            =   105
         TabIndex        =   294
         Top             =   4290
         Width           =   690
      End
      Begin VB.ComboBox cmb_Disp2 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   5
         ItemData        =   "FrmTcpServer.frx":0255
         Left            =   2520
         List            =   "FrmTcpServer.frx":0262
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   293
         ToolTipText     =   "���󺯰� �� �Ʒ��� ""�Ϲ�""��ư�� �����ּ���"
         Top             =   3915
         Width           =   615
      End
      Begin VB.ComboBox cmb_Disp1 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   5
         ItemData        =   "FrmTcpServer.frx":0272
         Left            =   2520
         List            =   "FrmTcpServer.frx":027F
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   292
         ToolTipText     =   "���󺯰� �� �Ʒ��� ""�Ϲ�""��ư�� �����ּ���"
         Top             =   3585
         Width           =   615
      End
      Begin VB.TextBox txt_Disp2 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   315
         Index           =   5
         Left            =   90
         TabIndex        =   291
         Text            =   "�����峻 ���� ����"
         Top             =   3915
         Width           =   2430
      End
      Begin VB.TextBox txt_Disp1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   315
         Index           =   5
         Left            =   90
         TabIndex        =   290
         Text            =   "�ϴ� ����..!!"
         Top             =   3585
         Width           =   2430
      End
      Begin VB.TextBox txt_DeviceIP 
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   5
         Left            =   90
         TabIndex        =   289
         Text            =   "192.168.0.211"
         ToolTipText     =   "�����Ʈ IP�ּҸ� �Է��ϼ���"
         Top             =   2550
         Width           =   1470
      End
      Begin VB.TextBox txt_GateName 
         Height          =   315
         Index           =   5
         Left            =   885
         TabIndex        =   288
         Text            =   "����"
         Top             =   600
         Width           =   1725
      End
      Begin VB.TextBox txt_DispIP 
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   5
         Left            =   1620
         TabIndex        =   287
         Text            =   "192.168.0.221"
         ToolTipText     =   "�����Ʈ IP�ּҸ� �Է��ϼ���"
         Top             =   2550
         Width           =   1470
      End
      Begin Threed.SSCommand cmd_DeviceReset 
         Height          =   510
         Index           =   5
         Left            =   330
         TabIndex        =   304
         ToolTipText     =   "���ܱ� ����� ����̽� �����մϴ�."
         Top             =   2910
         Width           =   915
         _Version        =   65536
         _ExtentX        =   1614
         _ExtentY        =   900
         _StockProps     =   78
         Caption         =   "����"
         ForeColor       =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "������� ExtraBold"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "FrmTcpServer.frx":028F
      End
      Begin Threed.SSCommand cmd_DispReset 
         Height          =   510
         Index           =   5
         Left            =   1920
         TabIndex        =   330
         ToolTipText     =   "������ �����մϴ�."
         Top             =   2910
         Width           =   915
         _Version        =   65536
         _ExtentX        =   1614
         _ExtentY        =   900
         _StockProps     =   78
         Caption         =   "����"
         ForeColor       =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "������� ExtraBold"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "FrmTcpServer.frx":05E0
      End
      Begin VB.Label lbl_BackCamera 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�Ĺ�"
         Height          =   210
         Index           =   5
         Left            =   270
         TabIndex        =   313
         Top             =   1350
         Width           =   540
      End
      Begin VB.Label lbl_PrintModel 
         BackColor       =   &H00E0E0E0&
         Caption         =   "����"
         Height          =   210
         Index           =   5
         Left            =   1305
         TabIndex        =   312
         Top             =   1410
         Width           =   360
      End
      Begin VB.Label lbl_PrintPort 
         BackColor       =   &H00E0E0E0&
         Caption         =   "��Ʈ"
         Height          =   210
         Index           =   5
         Left            =   1305
         TabIndex        =   311
         Top             =   1740
         Width           =   360
      End
      Begin VB.Label lbl_Guest 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�湮��"
         Height          =   210
         Index           =   5
         Left            =   270
         TabIndex        =   310
         Top             =   1740
         Width           =   540
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "��ġ"
         Height          =   210
         Index           =   26
         Left            =   270
         TabIndex        =   309
         Top             =   1020
         Width           =   840
      End
      Begin VB.Label lbl_DeviceIP 
         BackColor       =   &H00E0E0E0&
         Caption         =   "����̽� ������"
         Height          =   165
         Index           =   5
         Left            =   120
         TabIndex        =   308
         Top             =   2340
         Width           =   1360
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "����"
         Height          =   285
         Index           =   6
         Left            =   270
         TabIndex        =   307
         Top             =   300
         Width           =   495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   5
         X1              =   45
         X2              =   3120
         Y1              =   2145
         Y2              =   2145
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "��Ī"
         Height          =   210
         Index           =   25
         Left            =   270
         TabIndex        =   306
         Top             =   645
         Width           =   840
      End
      Begin VB.Label lbl_DispIP 
         BackColor       =   &H00E0E0E0&
         Caption         =   "������ ������"
         Height          =   165
         Index           =   5
         Left            =   1650
         TabIndex        =   305
         Top             =   2340
         Width           =   1365
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   " LANE5 "
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   5100
      Index           =   4
      Left            =   12900
      TabIndex        =   257
      Top             =   2910
      Width           =   3195
      Begin VB.CommandButton cmd_GateTestDown 
         Caption         =   "����"
         Height          =   330
         Index           =   4
         Left            =   870
         TabIndex        =   363
         Top             =   4680
         Width           =   690
      End
      Begin VB.ComboBox cmb_DispShiftSpeed 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   4
         ItemData        =   "FrmTcpServer.frx":0931
         Left            =   1710
         List            =   "FrmTcpServer.frx":0950
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   351
         ToolTipText     =   "������ �̵� �ӵ�����(���ڰ� ���� ���� ���� �̵��մϴ�), ���� �� �Ϲݹ�ư ��������"
         Top             =   4695
         Width           =   615
      End
      Begin VB.CommandButton cmd_NmlShift 
         Caption         =   "�̵�"
         Height          =   330
         Index           =   4
         Left            =   2415
         TabIndex        =   335
         ToolTipText     =   "���ι��� �Ϲ������� 6���ڱ��� �����մϴ�"
         Top             =   4680
         Width           =   690
      End
      Begin VB.CheckBox chk_BackCam_YN 
         Caption         =   "Check3"
         Height          =   225
         Index           =   4
         Left            =   885
         TabIndex        =   274
         Top             =   1350
         Width           =   195
      End
      Begin VB.ComboBox cmb_PrintModel 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         ItemData        =   "FrmTcpServer.frx":0978
         Left            =   1695
         List            =   "FrmTcpServer.frx":0982
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   273
         Top             =   1320
         Width           =   945
      End
      Begin VB.ComboBox cmb_PrintPort 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         ItemData        =   "FrmTcpServer.frx":0996
         Left            =   1695
         List            =   "FrmTcpServer.frx":09DF
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   272
         Top             =   1680
         Width           =   945
      End
      Begin VB.CheckBox chk_GuestYN 
         Caption         =   "Check3"
         Height          =   225
         Index           =   4
         Left            =   885
         TabIndex        =   271
         Top             =   1740
         Width           =   195
      End
      Begin VB.ComboBox CmbScreen 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         ItemData        =   "FrmTcpServer.frx":0A78
         Left            =   885
         List            =   "FrmTcpServer.frx":0A7A
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   270
         Top             =   930
         Width           =   1755
      End
      Begin VB.ComboBox cmb_Inout 
         Height          =   330
         Index           =   4
         ItemData        =   "FrmTcpServer.frx":0A7C
         Left            =   885
         List            =   "FrmTcpServer.frx":0A86
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   269
         Top             =   240
         Width           =   1725
      End
      Begin VB.CommandButton cmd_GateTest 
         Caption         =   "���ܱ�"
         Height          =   330
         Index           =   4
         Left            =   870
         TabIndex        =   268
         Top             =   4290
         Width           =   690
      End
      Begin VB.CommandButton cmd_EmgTest 
         Caption         =   "���"
         Height          =   330
         Index           =   4
         Left            =   1635
         TabIndex        =   267
         Top             =   4290
         Width           =   690
      End
      Begin VB.CommandButton cmd_NmlTest 
         Caption         =   "����"
         Height          =   330
         Index           =   4
         Left            =   2415
         TabIndex        =   266
         Top             =   4290
         Width           =   690
      End
      Begin VB.CommandButton cmd_CapTest 
         Caption         =   "ĸ��"
         Height          =   330
         Index           =   4
         Left            =   105
         TabIndex        =   265
         Top             =   4290
         Width           =   690
      End
      Begin VB.ComboBox cmb_Disp2 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   4
         ItemData        =   "FrmTcpServer.frx":0A96
         Left            =   2520
         List            =   "FrmTcpServer.frx":0AA3
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   264
         ToolTipText     =   "���󺯰� �� �Ʒ��� ""�Ϲ�""��ư�� �����ּ���"
         Top             =   3915
         Width           =   615
      End
      Begin VB.ComboBox cmb_Disp1 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   4
         ItemData        =   "FrmTcpServer.frx":0AB3
         Left            =   2520
         List            =   "FrmTcpServer.frx":0AC0
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   263
         ToolTipText     =   "���󺯰� �� �Ʒ��� ""�Ϲ�""��ư�� �����ּ���"
         Top             =   3585
         Width           =   615
      End
      Begin VB.TextBox txt_Disp2 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   315
         Index           =   4
         Left            =   90
         TabIndex        =   262
         Text            =   "�����峻 ���� ����"
         Top             =   3915
         Width           =   2430
      End
      Begin VB.TextBox txt_Disp1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   315
         Index           =   4
         Left            =   90
         TabIndex        =   261
         Text            =   "�ϴ� ����..!!"
         Top             =   3585
         Width           =   2430
      End
      Begin VB.TextBox txt_DeviceIP 
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   4
         Left            =   90
         TabIndex        =   260
         Text            =   "192.168.0.211"
         ToolTipText     =   "�����Ʈ IP�ּҸ� �Է��ϼ���"
         Top             =   2550
         Width           =   1470
      End
      Begin VB.TextBox txt_GateName 
         Height          =   315
         Index           =   4
         Left            =   885
         TabIndex        =   259
         Text            =   "����"
         Top             =   600
         Width           =   1725
      End
      Begin VB.TextBox txt_DispIP 
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   4
         Left            =   1620
         TabIndex        =   258
         Text            =   "192.168.0.221"
         ToolTipText     =   "�����Ʈ IP�ּҸ� �Է��ϼ���"
         Top             =   2550
         Width           =   1470
      End
      Begin Threed.SSCommand cmd_DeviceReset 
         Height          =   510
         Index           =   4
         Left            =   330
         TabIndex        =   275
         ToolTipText     =   "���ܱ� ����� ����̽� �����մϴ�."
         Top             =   2910
         Width           =   915
         _Version        =   65536
         _ExtentX        =   1614
         _ExtentY        =   900
         _StockProps     =   78
         Caption         =   "����"
         ForeColor       =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "������� ExtraBold"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "FrmTcpServer.frx":0AD0
      End
      Begin Threed.SSCommand cmd_DispReset 
         Height          =   510
         Index           =   4
         Left            =   1920
         TabIndex        =   329
         ToolTipText     =   "������ �����մϴ�."
         Top             =   2910
         Width           =   915
         _Version        =   65536
         _ExtentX        =   1614
         _ExtentY        =   900
         _StockProps     =   78
         Caption         =   "����"
         ForeColor       =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "������� ExtraBold"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "FrmTcpServer.frx":0E21
      End
      Begin VB.Label lbl_BackCamera 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�Ĺ�"
         Height          =   210
         Index           =   4
         Left            =   270
         TabIndex        =   284
         Top             =   1350
         Width           =   540
      End
      Begin VB.Label lbl_PrintModel 
         BackColor       =   &H00E0E0E0&
         Caption         =   "����"
         Height          =   210
         Index           =   4
         Left            =   1305
         TabIndex        =   283
         Top             =   1410
         Width           =   360
      End
      Begin VB.Label lbl_PrintPort 
         BackColor       =   &H00E0E0E0&
         Caption         =   "��Ʈ"
         Height          =   210
         Index           =   4
         Left            =   1305
         TabIndex        =   282
         Top             =   1740
         Width           =   360
      End
      Begin VB.Label lbl_Guest 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�湮��"
         Height          =   210
         Index           =   4
         Left            =   270
         TabIndex        =   281
         Top             =   1740
         Width           =   540
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "��ġ"
         Height          =   210
         Index           =   24
         Left            =   270
         TabIndex        =   280
         Top             =   1020
         Width           =   840
      End
      Begin VB.Label lbl_DeviceIP 
         BackColor       =   &H00E0E0E0&
         Caption         =   "����̽� ������"
         Height          =   165
         Index           =   4
         Left            =   120
         TabIndex        =   279
         Top             =   2340
         Width           =   1360
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "����"
         Height          =   285
         Index           =   5
         Left            =   270
         TabIndex        =   278
         Top             =   300
         Width           =   495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   4
         X1              =   45
         X2              =   3120
         Y1              =   2145
         Y2              =   2145
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "��Ī"
         Height          =   210
         Index           =   23
         Left            =   270
         TabIndex        =   277
         Top             =   645
         Width           =   840
      End
      Begin VB.Label lbl_DispIP 
         BackColor       =   &H00E0E0E0&
         Caption         =   "������ ������"
         Height          =   165
         Index           =   4
         Left            =   1650
         TabIndex        =   276
         Top             =   2340
         Width           =   1365
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   " LANE4 "
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   5100
      Index           =   3
      Left            =   9690
      TabIndex        =   228
      Top             =   2910
      Width           =   3195
      Begin VB.CommandButton cmd_GateTestDown 
         Caption         =   "����"
         Height          =   330
         Index           =   3
         Left            =   885
         TabIndex        =   367
         Top             =   4680
         Width           =   690
      End
      Begin VB.ComboBox cmb_DispShiftSpeed 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   3
         ItemData        =   "FrmTcpServer.frx":1172
         Left            =   1725
         List            =   "FrmTcpServer.frx":1191
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   350
         ToolTipText     =   "������ �̵� �ӵ�����(���ڰ� ���� ���� ���� �̵��մϴ�), ���� �� �Ϲݹ�ư ��������"
         Top             =   4695
         Width           =   615
      End
      Begin VB.CommandButton cmd_NmlShift 
         Caption         =   "�̵�"
         Height          =   330
         Index           =   3
         Left            =   2415
         TabIndex        =   334
         ToolTipText     =   "���ι��� �Ϲ������� 6���ڱ��� �����մϴ�"
         Top             =   4680
         Width           =   690
      End
      Begin VB.CheckBox chk_BackCam_YN 
         Caption         =   "Check3"
         Height          =   225
         Index           =   3
         Left            =   885
         TabIndex        =   245
         Top             =   1350
         Width           =   195
      End
      Begin VB.ComboBox cmb_PrintModel 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         ItemData        =   "FrmTcpServer.frx":11B9
         Left            =   1695
         List            =   "FrmTcpServer.frx":11C3
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   244
         Top             =   1320
         Width           =   945
      End
      Begin VB.ComboBox cmb_PrintPort 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         ItemData        =   "FrmTcpServer.frx":11D7
         Left            =   1695
         List            =   "FrmTcpServer.frx":1220
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   243
         Top             =   1680
         Width           =   945
      End
      Begin VB.CheckBox chk_GuestYN 
         Caption         =   "Check3"
         Height          =   225
         Index           =   3
         Left            =   885
         TabIndex        =   242
         Top             =   1740
         Width           =   195
      End
      Begin VB.ComboBox CmbScreen 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         ItemData        =   "FrmTcpServer.frx":12B9
         Left            =   885
         List            =   "FrmTcpServer.frx":12BB
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   241
         Top             =   930
         Width           =   1755
      End
      Begin VB.ComboBox cmb_Inout 
         Height          =   330
         Index           =   3
         ItemData        =   "FrmTcpServer.frx":12BD
         Left            =   885
         List            =   "FrmTcpServer.frx":12C7
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   240
         Top             =   240
         Width           =   1725
      End
      Begin VB.CommandButton cmd_GateTest 
         Caption         =   "���ܱ�"
         Height          =   330
         Index           =   3
         Left            =   885
         TabIndex        =   239
         Top             =   4290
         Width           =   690
      End
      Begin VB.CommandButton cmd_EmgTest 
         Caption         =   "���"
         Height          =   330
         Index           =   3
         Left            =   1650
         TabIndex        =   238
         Top             =   4290
         Width           =   690
      End
      Begin VB.CommandButton cmd_NmlTest 
         Caption         =   "����"
         Height          =   330
         Index           =   3
         Left            =   2415
         TabIndex        =   237
         Top             =   4290
         Width           =   690
      End
      Begin VB.CommandButton cmd_CapTest 
         Caption         =   "ĸ��"
         Height          =   330
         Index           =   3
         Left            =   105
         TabIndex        =   236
         Top             =   4290
         Width           =   690
      End
      Begin VB.ComboBox cmb_Disp2 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   3
         ItemData        =   "FrmTcpServer.frx":12D7
         Left            =   2520
         List            =   "FrmTcpServer.frx":12E4
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   235
         ToolTipText     =   "���󺯰� �� �Ʒ��� ""�Ϲ�""��ư�� �����ּ���"
         Top             =   3915
         Width           =   615
      End
      Begin VB.ComboBox cmb_Disp1 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   3
         ItemData        =   "FrmTcpServer.frx":12F4
         Left            =   2520
         List            =   "FrmTcpServer.frx":1301
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   234
         ToolTipText     =   "���󺯰� �� �Ʒ��� ""�Ϲ�""��ư�� �����ּ���"
         Top             =   3585
         Width           =   615
      End
      Begin VB.TextBox txt_Disp2 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   315
         Index           =   3
         Left            =   90
         TabIndex        =   233
         Text            =   "�����峻 ���� ����"
         Top             =   3915
         Width           =   2430
      End
      Begin VB.TextBox txt_Disp1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   315
         Index           =   3
         Left            =   90
         TabIndex        =   232
         Text            =   "�ϴ� ����..!!"
         Top             =   3585
         Width           =   2430
      End
      Begin VB.TextBox txt_DeviceIP 
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   3
         Left            =   90
         TabIndex        =   231
         Text            =   "192.168.0.211"
         ToolTipText     =   "�����Ʈ IP�ּҸ� �Է��ϼ���"
         Top             =   2550
         Width           =   1470
      End
      Begin VB.TextBox txt_GateName 
         Height          =   315
         Index           =   3
         Left            =   885
         TabIndex        =   230
         Text            =   "����"
         Top             =   600
         Width           =   1725
      End
      Begin VB.TextBox txt_DispIP 
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   3
         Left            =   1620
         TabIndex        =   229
         Text            =   "192.168.0.221"
         ToolTipText     =   "�����Ʈ IP�ּҸ� �Է��ϼ���"
         Top             =   2550
         Width           =   1470
      End
      Begin Threed.SSCommand cmd_DeviceReset 
         Height          =   510
         Index           =   3
         Left            =   330
         TabIndex        =   246
         ToolTipText     =   "���ܱ� ����� ����̽� �����մϴ�."
         Top             =   2910
         Width           =   915
         _Version        =   65536
         _ExtentX        =   1614
         _ExtentY        =   900
         _StockProps     =   78
         Caption         =   "����"
         ForeColor       =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "������� ExtraBold"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "FrmTcpServer.frx":1311
      End
      Begin Threed.SSCommand cmd_DispReset 
         Height          =   510
         Index           =   3
         Left            =   1905
         TabIndex        =   328
         ToolTipText     =   "������ �����մϴ�."
         Top             =   2910
         Width           =   915
         _Version        =   65536
         _ExtentX        =   1614
         _ExtentY        =   900
         _StockProps     =   78
         Caption         =   "����"
         ForeColor       =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "������� ExtraBold"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "FrmTcpServer.frx":1662
      End
      Begin VB.Label lbl_BackCamera 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�Ĺ�"
         Height          =   210
         Index           =   3
         Left            =   270
         TabIndex        =   255
         Top             =   1350
         Width           =   540
      End
      Begin VB.Label lbl_PrintModel 
         BackColor       =   &H00E0E0E0&
         Caption         =   "����"
         Height          =   210
         Index           =   3
         Left            =   1305
         TabIndex        =   254
         Top             =   1410
         Width           =   360
      End
      Begin VB.Label lbl_PrintPort 
         BackColor       =   &H00E0E0E0&
         Caption         =   "��Ʈ"
         Height          =   210
         Index           =   3
         Left            =   1305
         TabIndex        =   253
         Top             =   1740
         Width           =   360
      End
      Begin VB.Label lbl_Guest 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�湮��"
         Height          =   210
         Index           =   3
         Left            =   270
         TabIndex        =   252
         Top             =   1740
         Width           =   540
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "��ġ"
         Height          =   210
         Index           =   19
         Left            =   270
         TabIndex        =   251
         Top             =   1020
         Width           =   840
      End
      Begin VB.Label lbl_DeviceIP 
         BackColor       =   &H00E0E0E0&
         Caption         =   "����̽� ������"
         Height          =   165
         Index           =   3
         Left            =   120
         TabIndex        =   250
         Top             =   2340
         Width           =   1360
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "����"
         Height          =   285
         Index           =   3
         Left            =   270
         TabIndex        =   249
         Top             =   300
         Width           =   495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   3
         X1              =   45
         X2              =   3120
         Y1              =   2145
         Y2              =   2145
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "��Ī"
         Height          =   210
         Index           =   16
         Left            =   270
         TabIndex        =   248
         Top             =   645
         Width           =   840
      End
      Begin VB.Label lbl_DispIP 
         BackColor       =   &H00E0E0E0&
         Caption         =   "������ ������"
         Height          =   165
         Index           =   3
         Left            =   1635
         TabIndex        =   247
         Top             =   2340
         Width           =   1365
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   " LANE3 "
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   5100
      Index           =   2
      Left            =   6480
      TabIndex        =   199
      Top             =   2910
      Width           =   3195
      Begin VB.CommandButton cmd_GateTestDown 
         Caption         =   "����"
         Height          =   330
         Index           =   2
         Left            =   870
         TabIndex        =   366
         Top             =   4680
         Width           =   690
      End
      Begin VB.ComboBox cmb_DispShiftSpeed 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   2
         ItemData        =   "FrmTcpServer.frx":19B3
         Left            =   1725
         List            =   "FrmTcpServer.frx":19D2
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   349
         ToolTipText     =   "������ �̵� �ӵ�����(���ڰ� ���� ���� ���� �̵��մϴ�), ���� �� �Ϲݹ�ư ��������"
         Top             =   4695
         Width           =   615
      End
      Begin VB.CommandButton cmd_NmlShift 
         Caption         =   "�̵�"
         Height          =   330
         Index           =   2
         Left            =   2415
         TabIndex        =   333
         ToolTipText     =   "���ι��� �Ϲ������� 6���ڱ��� �����մϴ�"
         Top             =   4680
         Width           =   690
      End
      Begin VB.CheckBox chk_BackCam_YN 
         Caption         =   "Check3"
         Height          =   225
         Index           =   2
         Left            =   885
         TabIndex        =   216
         Top             =   1350
         Width           =   195
      End
      Begin VB.ComboBox cmb_PrintModel 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         ItemData        =   "FrmTcpServer.frx":19FA
         Left            =   1695
         List            =   "FrmTcpServer.frx":1A04
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   215
         Top             =   1320
         Width           =   945
      End
      Begin VB.ComboBox cmb_PrintPort 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         ItemData        =   "FrmTcpServer.frx":1A18
         Left            =   1695
         List            =   "FrmTcpServer.frx":1A61
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   214
         Top             =   1680
         Width           =   945
      End
      Begin VB.CheckBox chk_GuestYN 
         Caption         =   "Check3"
         Height          =   225
         Index           =   2
         Left            =   885
         TabIndex        =   213
         Top             =   1740
         Width           =   195
      End
      Begin VB.ComboBox CmbScreen 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         ItemData        =   "FrmTcpServer.frx":1AFA
         Left            =   885
         List            =   "FrmTcpServer.frx":1AFC
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   212
         Top             =   930
         Width           =   1755
      End
      Begin VB.ComboBox cmb_Inout 
         Height          =   330
         Index           =   2
         ItemData        =   "FrmTcpServer.frx":1AFE
         Left            =   885
         List            =   "FrmTcpServer.frx":1B08
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   211
         Top             =   240
         Width           =   1725
      End
      Begin VB.CommandButton cmd_GateTest 
         Caption         =   "���ܱ�"
         Height          =   330
         Index           =   2
         Left            =   870
         TabIndex        =   210
         Top             =   4290
         Width           =   690
      End
      Begin VB.CommandButton cmd_EmgTest 
         Caption         =   "���"
         Height          =   330
         Index           =   2
         Left            =   1650
         TabIndex        =   209
         Top             =   4290
         Width           =   690
      End
      Begin VB.CommandButton cmd_NmlTest 
         Caption         =   "����"
         Height          =   330
         Index           =   2
         Left            =   2415
         TabIndex        =   208
         Top             =   4290
         Width           =   690
      End
      Begin VB.CommandButton cmd_CapTest 
         Caption         =   "ĸ��"
         Height          =   330
         Index           =   2
         Left            =   105
         TabIndex        =   207
         Top             =   4290
         Width           =   690
      End
      Begin VB.ComboBox cmb_Disp2 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   2
         ItemData        =   "FrmTcpServer.frx":1B18
         Left            =   2520
         List            =   "FrmTcpServer.frx":1B25
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   206
         ToolTipText     =   "���󺯰� �� �Ʒ��� ""�Ϲ�""��ư�� �����ּ���"
         Top             =   3915
         Width           =   615
      End
      Begin VB.ComboBox cmb_Disp1 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   2
         ItemData        =   "FrmTcpServer.frx":1B35
         Left            =   2520
         List            =   "FrmTcpServer.frx":1B42
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   205
         ToolTipText     =   "���󺯰� �� �Ʒ��� ""�Ϲ�""��ư�� �����ּ���"
         Top             =   3585
         Width           =   615
      End
      Begin VB.TextBox txt_Disp2 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   315
         Index           =   2
         Left            =   90
         TabIndex        =   204
         Text            =   "�����峻 ���� ����"
         Top             =   3915
         Width           =   2430
      End
      Begin VB.TextBox txt_Disp1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   315
         Index           =   2
         Left            =   90
         TabIndex        =   203
         Text            =   "�ϴ� ����..!!"
         Top             =   3585
         Width           =   2430
      End
      Begin VB.TextBox txt_DeviceIP 
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   2
         Left            =   90
         TabIndex        =   202
         Text            =   "192.168.0.211"
         ToolTipText     =   "�����Ʈ IP�ּҸ� �Է��ϼ���"
         Top             =   2550
         Width           =   1470
      End
      Begin VB.TextBox txt_GateName 
         Height          =   315
         Index           =   2
         Left            =   885
         TabIndex        =   201
         Text            =   "����"
         Top             =   600
         Width           =   1725
      End
      Begin VB.TextBox txt_DispIP 
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   2
         Left            =   1620
         TabIndex        =   200
         Text            =   "192.168.0.221"
         ToolTipText     =   "�����Ʈ IP�ּҸ� �Է��ϼ���"
         Top             =   2550
         Width           =   1470
      End
      Begin Threed.SSCommand cmd_DeviceReset 
         Height          =   510
         Index           =   2
         Left            =   330
         TabIndex        =   217
         ToolTipText     =   "���ܱ� ����� ����̽� �����մϴ�."
         Top             =   2910
         Width           =   915
         _Version        =   65536
         _ExtentX        =   1614
         _ExtentY        =   900
         _StockProps     =   78
         Caption         =   "����"
         ForeColor       =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "������� ExtraBold"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "FrmTcpServer.frx":1B52
      End
      Begin Threed.SSCommand cmd_DispReset 
         Height          =   510
         Index           =   2
         Left            =   1920
         TabIndex        =   327
         ToolTipText     =   "������ �����մϴ�."
         Top             =   2910
         Width           =   915
         _Version        =   65536
         _ExtentX        =   1614
         _ExtentY        =   900
         _StockProps     =   78
         Caption         =   "����"
         ForeColor       =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "������� ExtraBold"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "FrmTcpServer.frx":1EA3
      End
      Begin VB.Label lbl_BackCamera 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�Ĺ�"
         Height          =   210
         Index           =   2
         Left            =   270
         TabIndex        =   226
         Top             =   1350
         Width           =   540
      End
      Begin VB.Label lbl_PrintModel 
         BackColor       =   &H00E0E0E0&
         Caption         =   "����"
         Height          =   210
         Index           =   2
         Left            =   1305
         TabIndex        =   225
         Top             =   1410
         Width           =   360
      End
      Begin VB.Label lbl_PrintPort 
         BackColor       =   &H00E0E0E0&
         Caption         =   "��Ʈ"
         Height          =   210
         Index           =   2
         Left            =   1305
         TabIndex        =   224
         Top             =   1740
         Width           =   360
      End
      Begin VB.Label lbl_Guest 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�湮��"
         Height          =   210
         Index           =   2
         Left            =   270
         TabIndex        =   223
         Top             =   1740
         Width           =   540
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "��ġ"
         Height          =   210
         Index           =   14
         Left            =   270
         TabIndex        =   222
         Top             =   1020
         Width           =   840
      End
      Begin VB.Label lbl_DeviceIP 
         BackColor       =   &H00E0E0E0&
         Caption         =   "����̽� ������"
         Height          =   165
         Index           =   2
         Left            =   120
         TabIndex        =   221
         Top             =   2340
         Width           =   1360
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "����"
         Height          =   285
         Index           =   2
         Left            =   270
         TabIndex        =   220
         Top             =   300
         Width           =   495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   2
         X1              =   45
         X2              =   3120
         Y1              =   2145
         Y2              =   2145
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "��Ī"
         Height          =   210
         Index           =   11
         Left            =   270
         TabIndex        =   219
         Top             =   645
         Width           =   840
      End
      Begin VB.Label lbl_DispIP 
         BackColor       =   &H00E0E0E0&
         Caption         =   "������ ������"
         Height          =   165
         Index           =   2
         Left            =   1650
         TabIndex        =   218
         Top             =   2340
         Width           =   1365
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   " LANE2 "
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   5100
      Index           =   1
      Left            =   3270
      TabIndex        =   170
      Top             =   2910
      Width           =   3195
      Begin VB.CommandButton cmd_GateTestDown 
         Caption         =   "����"
         Height          =   330
         Index           =   1
         Left            =   885
         TabIndex        =   365
         Top             =   4680
         Width           =   690
      End
      Begin VB.ComboBox cmb_DispShiftSpeed 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   1
         ItemData        =   "FrmTcpServer.frx":21F4
         Left            =   1725
         List            =   "FrmTcpServer.frx":2213
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   348
         ToolTipText     =   "������ �̵� �ӵ�����(���ڰ� ���� ���� ���� �̵��մϴ�), ���� �� �Ϲݹ�ư ��������"
         Top             =   4695
         Width           =   615
      End
      Begin VB.CommandButton cmd_NmlShift 
         Caption         =   "�̵�"
         Height          =   330
         Index           =   1
         Left            =   2415
         TabIndex        =   332
         ToolTipText     =   "���ι��� �Ϲ������� 6���ڱ��� �����մϴ�"
         Top             =   4680
         Width           =   690
      End
      Begin VB.CheckBox chk_BackCam_YN 
         Caption         =   "Check3"
         Height          =   225
         Index           =   1
         Left            =   885
         TabIndex        =   187
         Top             =   1350
         Width           =   195
      End
      Begin VB.ComboBox cmb_PrintModel 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         ItemData        =   "FrmTcpServer.frx":223B
         Left            =   1695
         List            =   "FrmTcpServer.frx":2245
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   186
         Top             =   1320
         Width           =   945
      End
      Begin VB.ComboBox cmb_PrintPort 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         ItemData        =   "FrmTcpServer.frx":2259
         Left            =   1695
         List            =   "FrmTcpServer.frx":22A2
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   185
         Top             =   1680
         Width           =   945
      End
      Begin VB.CheckBox chk_GuestYN 
         Caption         =   "Check3"
         Height          =   225
         Index           =   1
         Left            =   885
         TabIndex        =   184
         Top             =   1740
         Width           =   195
      End
      Begin VB.ComboBox CmbScreen 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         ItemData        =   "FrmTcpServer.frx":233B
         Left            =   885
         List            =   "FrmTcpServer.frx":233D
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   183
         Top             =   930
         Width           =   1755
      End
      Begin VB.ComboBox cmb_Inout 
         Height          =   330
         Index           =   1
         ItemData        =   "FrmTcpServer.frx":233F
         Left            =   885
         List            =   "FrmTcpServer.frx":2349
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   182
         Top             =   240
         Width           =   1725
      End
      Begin VB.CommandButton cmd_GateTest 
         Caption         =   "���ܱ�"
         Height          =   330
         Index           =   1
         Left            =   885
         TabIndex        =   181
         Top             =   4290
         Width           =   690
      End
      Begin VB.CommandButton cmd_EmgTest 
         Caption         =   "���"
         Height          =   330
         Index           =   1
         Left            =   1650
         TabIndex        =   180
         Top             =   4290
         Width           =   690
      End
      Begin VB.CommandButton cmd_NmlTest 
         Caption         =   "����"
         Height          =   330
         Index           =   1
         Left            =   2415
         TabIndex        =   179
         Top             =   4290
         Width           =   690
      End
      Begin VB.CommandButton cmd_CapTest 
         Caption         =   "ĸ��"
         Height          =   330
         Index           =   1
         Left            =   105
         TabIndex        =   178
         Top             =   4290
         Width           =   690
      End
      Begin VB.ComboBox cmb_Disp2 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   1
         ItemData        =   "FrmTcpServer.frx":2359
         Left            =   2520
         List            =   "FrmTcpServer.frx":2366
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   177
         ToolTipText     =   "���󺯰� �� �Ʒ��� ""�Ϲ�""��ư�� �����ּ���"
         Top             =   3915
         Width           =   615
      End
      Begin VB.ComboBox cmb_Disp1 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   1
         ItemData        =   "FrmTcpServer.frx":2376
         Left            =   2520
         List            =   "FrmTcpServer.frx":2383
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   176
         ToolTipText     =   "���󺯰� �� �Ʒ��� ""�Ϲ�""��ư�� �����ּ���"
         Top             =   3585
         Width           =   615
      End
      Begin VB.TextBox txt_Disp2 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   315
         Index           =   1
         Left            =   90
         TabIndex        =   175
         Text            =   "�����峻 ���� ����"
         Top             =   3915
         Width           =   2430
      End
      Begin VB.TextBox txt_Disp1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   315
         Index           =   1
         Left            =   90
         TabIndex        =   174
         Text            =   "�ϴ� ����..!!"
         Top             =   3585
         Width           =   2430
      End
      Begin VB.TextBox txt_DeviceIP 
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   1
         Left            =   90
         TabIndex        =   173
         Text            =   "192.168.0.211"
         ToolTipText     =   "�����Ʈ IP�ּҸ� �Է��ϼ���"
         Top             =   2550
         Width           =   1470
      End
      Begin VB.TextBox txt_GateName 
         Height          =   315
         Index           =   1
         Left            =   885
         TabIndex        =   172
         Text            =   "����"
         Top             =   600
         Width           =   1725
      End
      Begin VB.TextBox txt_DispIP 
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   1
         Left            =   1620
         TabIndex        =   171
         Text            =   "192.168.0.221"
         ToolTipText     =   "�����Ʈ IP�ּҸ� �Է��ϼ���"
         Top             =   2550
         Width           =   1470
      End
      Begin Threed.SSCommand cmd_DeviceReset 
         Height          =   510
         Index           =   1
         Left            =   330
         TabIndex        =   188
         ToolTipText     =   "���ܱ� ����� ����̽� �����մϴ�."
         Top             =   2910
         Width           =   915
         _Version        =   65536
         _ExtentX        =   1614
         _ExtentY        =   900
         _StockProps     =   78
         Caption         =   "����"
         ForeColor       =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "������� ExtraBold"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "FrmTcpServer.frx":2393
      End
      Begin Threed.SSCommand cmd_DispReset 
         Height          =   510
         Index           =   1
         Left            =   1920
         TabIndex        =   326
         ToolTipText     =   "������ �����մϴ�."
         Top             =   2910
         Width           =   915
         _Version        =   65536
         _ExtentX        =   1614
         _ExtentY        =   900
         _StockProps     =   78
         Caption         =   "����"
         ForeColor       =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "������� ExtraBold"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "FrmTcpServer.frx":26E4
      End
      Begin VB.Label lbl_BackCamera 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�Ĺ�"
         Height          =   210
         Index           =   1
         Left            =   270
         TabIndex        =   197
         Top             =   1350
         Width           =   540
      End
      Begin VB.Label lbl_PrintModel 
         BackColor       =   &H00E0E0E0&
         Caption         =   "����"
         Height          =   210
         Index           =   1
         Left            =   1305
         TabIndex        =   196
         Top             =   1410
         Width           =   360
      End
      Begin VB.Label lbl_PrintPort 
         BackColor       =   &H00E0E0E0&
         Caption         =   "��Ʈ"
         Height          =   210
         Index           =   1
         Left            =   1305
         TabIndex        =   195
         Top             =   1740
         Width           =   360
      End
      Begin VB.Label lbl_Guest 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�湮��"
         Height          =   210
         Index           =   1
         Left            =   270
         TabIndex        =   194
         Top             =   1740
         Width           =   540
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "��ġ"
         Height          =   210
         Index           =   9
         Left            =   270
         TabIndex        =   193
         Top             =   1020
         Width           =   840
      End
      Begin VB.Label lbl_DeviceIP 
         BackColor       =   &H00E0E0E0&
         Caption         =   "����̽� ������"
         Height          =   165
         Index           =   1
         Left            =   120
         TabIndex        =   192
         Top             =   2340
         Width           =   1360
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "����"
         Height          =   285
         Index           =   1
         Left            =   270
         TabIndex        =   191
         Top             =   300
         Width           =   495
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   1
         X1              =   45
         X2              =   3120
         Y1              =   2145
         Y2              =   2145
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "��Ī"
         Height          =   210
         Index           =   6
         Left            =   270
         TabIndex        =   190
         Top             =   645
         Width           =   840
      End
      Begin VB.Label lbl_DispIP 
         BackColor       =   &H00E0E0E0&
         Caption         =   "������ ������"
         Height          =   165
         Index           =   1
         Left            =   1650
         TabIndex        =   189
         Top             =   2340
         Width           =   1365
      End
   End
   Begin VB.CheckBox chk_UseYN 
      BackColor       =   &H00E0E0E0&
      Caption         =   "���"
      Height          =   210
      Index           =   0
      Left            =   2550
      TabIndex        =   114
      Top             =   2775
      Width           =   630
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   20325
      Top             =   -60
   End
   Begin VB.Frame frameLocalInfo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "�� ��ǻ��"
      ForeColor       =   &H00000000&
      Height          =   900
      Left            =   60
      TabIndex        =   145
      Top             =   1635
      Width           =   1605
      Begin VB.TextBox txtIP 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   375
         Left            =   105
         Locked          =   -1  'True
         TabIndex        =   146
         Text            =   "255.255.255.255"
         Top             =   375
         Width           =   1380
      End
   End
   Begin VB.ListBox ListData 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   2580
      Left            =   30
      TabIndex        =   144
      Top             =   8655
      Width           =   19290
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   210
      Left            =   -12780
      TabIndex        =   143
      Top             =   4500
      Width           =   75
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "�����ͼ���"
      ForeColor       =   &H00FF0000&
      Height          =   900
      Index           =   0
      Left            =   11685
      TabIndex        =   140
      Top             =   1635
      Width           =   1755
      Begin VB.TextBox TxtSvrPort 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   585
         TabIndex        =   141
         Text            =   "10000"
         ToolTipText     =   "ȣ��Ʈpc�� �pc�� �и��� ���  ""���"" üũ �ϼ���. �pc���� ������ ��Ʈ��ȣ �Դϴ�."
         Top             =   390
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "��Ʈ"
         Height          =   255
         Left            =   135
         TabIndex        =   142
         Top             =   450
         Visible         =   0   'False
         Width           =   435
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   " LANE1 "
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   5100
      Index           =   0
      Left            =   60
      TabIndex        =   115
      Top             =   2910
      Width           =   3195
      Begin VB.CommandButton cmd_GateTestDown 
         Caption         =   "����"
         Height          =   330
         Index           =   0
         Left            =   885
         TabIndex        =   364
         Top             =   4680
         Width           =   690
      End
      Begin VB.ComboBox cmb_DispShiftSpeed 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         ItemData        =   "FrmTcpServer.frx":2A35
         Left            =   1725
         List            =   "FrmTcpServer.frx":2A54
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   347
         ToolTipText     =   "������ �̵� �ӵ�����(���ڰ� ���� ���� ���� �̵��մϴ�), ���� �� �Ϲݹ�ư ��������"
         Top             =   4680
         Width           =   615
      End
      Begin VB.CommandButton cmd_NmlShift 
         Caption         =   "�̵�"
         Height          =   330
         Index           =   0
         Left            =   2415
         TabIndex        =   331
         ToolTipText     =   "���ι��� �Ϲ������� 6���ڱ��� �����մϴ�"
         Top             =   4680
         Width           =   690
      End
      Begin VB.TextBox txt_DispIP 
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   0
         Left            =   1620
         TabIndex        =   166
         Text            =   "192.168.0.221"
         ToolTipText     =   "�����Ʈ IP�ּҸ� �Է��ϼ���"
         Top             =   2550
         Width           =   1470
      End
      Begin VB.TextBox txt_GateName 
         Height          =   315
         Index           =   0
         Left            =   885
         TabIndex        =   131
         Text            =   "����"
         Top             =   600
         Width           =   1725
      End
      Begin VB.TextBox txt_DeviceIP 
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   0
         Left            =   90
         TabIndex        =   130
         Text            =   "192.168.0.211"
         ToolTipText     =   "�����Ʈ IP�ּҸ� �Է��ϼ���"
         Top             =   2550
         Width           =   1470
      End
      Begin VB.TextBox txt_Disp1 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   315
         Index           =   0
         Left            =   90
         TabIndex        =   129
         Text            =   "�ϴ� ����..!!"
         Top             =   3585
         Width           =   2430
      End
      Begin VB.TextBox txt_Disp2 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   315
         Index           =   0
         Left            =   90
         TabIndex        =   128
         Text            =   "�����峻 ���� ����"
         Top             =   3915
         Width           =   2430
      End
      Begin VB.ComboBox cmb_Disp1 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         ItemData        =   "FrmTcpServer.frx":2A7C
         Left            =   2520
         List            =   "FrmTcpServer.frx":2A95
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   127
         ToolTipText     =   "���󺯰� �� �Ʒ��� ""�Ϲ�""��ư�� �����ּ���"
         Top             =   3585
         Width           =   615
      End
      Begin VB.ComboBox cmb_Disp2 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         ItemData        =   "FrmTcpServer.frx":2AB5
         Left            =   2520
         List            =   "FrmTcpServer.frx":2ACE
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   126
         ToolTipText     =   "���󺯰� �� �Ʒ��� ""�Ϲ�""��ư�� �����ּ���"
         Top             =   3915
         Width           =   615
      End
      Begin VB.CommandButton cmd_CapTest 
         Caption         =   "ĸ��"
         Height          =   330
         Index           =   0
         Left            =   105
         TabIndex        =   125
         Top             =   4290
         Width           =   690
      End
      Begin VB.CommandButton cmd_NmlTest 
         Caption         =   "����"
         Height          =   330
         Index           =   0
         Left            =   2415
         TabIndex        =   124
         Top             =   4290
         Width           =   690
      End
      Begin VB.CommandButton cmd_EmgTest 
         Caption         =   "���"
         Height          =   330
         Index           =   0
         Left            =   1650
         TabIndex        =   123
         Top             =   4290
         Width           =   690
      End
      Begin VB.CommandButton cmd_GateTest 
         Caption         =   "���ܱ�"
         Height          =   330
         Index           =   0
         Left            =   885
         TabIndex        =   122
         Top             =   4290
         Width           =   690
      End
      Begin VB.ComboBox cmb_Inout 
         Height          =   330
         Index           =   0
         ItemData        =   "FrmTcpServer.frx":2AEE
         Left            =   885
         List            =   "FrmTcpServer.frx":2AF8
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   121
         Top             =   240
         Width           =   1725
      End
      Begin VB.ComboBox CmbScreen 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         ItemData        =   "FrmTcpServer.frx":2B08
         Left            =   885
         List            =   "FrmTcpServer.frx":2B0A
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   120
         Top             =   930
         Width           =   1755
      End
      Begin VB.CheckBox chk_GuestYN 
         Caption         =   "Check3"
         Height          =   225
         Index           =   0
         Left            =   885
         TabIndex        =   119
         Top             =   1740
         Width           =   195
      End
      Begin VB.ComboBox cmb_PrintPort 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         ItemData        =   "FrmTcpServer.frx":2B0C
         Left            =   1695
         List            =   "FrmTcpServer.frx":2B55
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   118
         Top             =   1680
         Width           =   945
      End
      Begin VB.ComboBox cmb_PrintModel 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         ItemData        =   "FrmTcpServer.frx":2BEE
         Left            =   1695
         List            =   "FrmTcpServer.frx":2BF8
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   117
         Top             =   1320
         Width           =   945
      End
      Begin VB.CheckBox chk_BackCam_YN 
         Caption         =   "Check3"
         Height          =   225
         Index           =   0
         Left            =   885
         TabIndex        =   116
         Top             =   1350
         Width           =   195
      End
      Begin Threed.SSCommand cmd_DeviceReset 
         Height          =   510
         Index           =   0
         Left            =   330
         TabIndex        =   168
         ToolTipText     =   "���ܱ� ����� ����̽� �����մϴ�."
         Top             =   2910
         Width           =   915
         _Version        =   65536
         _ExtentX        =   1614
         _ExtentY        =   900
         _StockProps     =   78
         Caption         =   "����"
         ForeColor       =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "������� ExtraBold"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "FrmTcpServer.frx":2C0C
      End
      Begin Threed.SSCommand cmd_DispReset 
         Height          =   510
         Index           =   0
         Left            =   1920
         TabIndex        =   325
         ToolTipText     =   "������ �����մϴ�."
         Top             =   2910
         Width           =   915
         _Version        =   65536
         _ExtentX        =   1614
         _ExtentY        =   900
         _StockProps     =   78
         Caption         =   "����"
         ForeColor       =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "������� ExtraBold"
            Size            =   11.25
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "FrmTcpServer.frx":2F5D
      End
      Begin VB.Label lbl_DispIP 
         BackColor       =   &H00E0E0E0&
         Caption         =   "������ ������"
         Height          =   165
         Index           =   0
         Left            =   1650
         TabIndex        =   167
         Top             =   2340
         Width           =   1360
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "��Ī"
         Height          =   210
         Index           =   2
         Left            =   270
         TabIndex        =   139
         Top             =   645
         Width           =   840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   45
         X2              =   3120
         Y1              =   2145
         Y2              =   2145
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "����"
         Height          =   285
         Index           =   0
         Left            =   270
         TabIndex        =   138
         Top             =   300
         Width           =   495
      End
      Begin VB.Label lbl_DeviceIP 
         BackColor       =   &H00E0E0E0&
         Caption         =   "����̽� ������"
         Height          =   165
         Index           =   0
         Left            =   120
         TabIndex        =   137
         Top             =   2340
         Width           =   1360
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "��ġ"
         Height          =   210
         Index           =   4
         Left            =   270
         TabIndex        =   136
         Top             =   1020
         Width           =   840
      End
      Begin VB.Label lbl_Guest 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�湮��"
         Height          =   210
         Index           =   0
         Left            =   270
         TabIndex        =   135
         Top             =   1740
         Width           =   540
      End
      Begin VB.Label lbl_PrintPort 
         BackColor       =   &H00E0E0E0&
         Caption         =   "��Ʈ"
         Height          =   210
         Index           =   0
         Left            =   1305
         TabIndex        =   134
         Top             =   1740
         Width           =   360
      End
      Begin VB.Label lbl_PrintModel 
         BackColor       =   &H00E0E0E0&
         Caption         =   "����"
         Height          =   210
         Index           =   0
         Left            =   1305
         TabIndex        =   133
         Top             =   1410
         Width           =   360
      End
      Begin VB.Label lbl_BackCamera 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�Ĺ�"
         Height          =   210
         Index           =   0
         Left            =   270
         TabIndex        =   132
         Top             =   1350
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "�����ͼ۽�"
      ForeColor       =   &H00FF0000&
      Height          =   900
      Index           =   1
      Left            =   13485
      TabIndex        =   109
      Top             =   1635
      Width           =   1650
      Begin VB.TextBox TxtSvrIp 
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   111
         Text            =   "255.255.255.255"
         ToolTipText     =   "����pc�� �pc�� �и��� ���  ""���"" üũ �ϼ���. ����pc ip�ּҸ� ��������"
         Top             =   480
         Width           =   1395
      End
      Begin VB.TextBox TxtSvrPort 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   1545
         TabIndex        =   110
         Text            =   "10000"
         ToolTipText     =   "ȣ��Ʈpc�� �pc�� �и��� ���  ""���"" üũ �ϼ���. ����pc�� �۽��� ��Ʈ��ȣ �Դϴ�."
         Top             =   480
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "������"
         Height          =   255
         Left            =   165
         TabIndex        =   113
         Top             =   270
         Width           =   885
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "��Ʈ"
         Height          =   225
         Left            =   1575
         TabIndex        =   112
         Top             =   270
         Visible         =   0   'False
         Width           =   675
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   " LANE5 "
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   5730
      Index           =   44
      Left            =   20550
      TabIndex        =   81
      Top             =   6600
      Visible         =   0   'False
      Width           =   3285
      Begin VB.CommandButton cmd_GateTest 
         Caption         =   "���ܱ�"
         Height          =   330
         Index           =   44
         Left            =   865
         TabIndex        =   100
         Top             =   5280
         Width           =   690
      End
      Begin VB.CommandButton cmd_EmgTest 
         Caption         =   "���"
         Height          =   330
         Index           =   44
         Left            =   1640
         TabIndex        =   99
         Top             =   5280
         Width           =   690
      End
      Begin VB.CommandButton cmd_NmlTest 
         Caption         =   "�Ϲ�"
         Height          =   330
         Index           =   44
         Left            =   2415
         TabIndex        =   98
         Top             =   5280
         Width           =   690
      End
      Begin VB.CommandButton cmd_CapTest 
         Caption         =   "ĸ��"
         Height          =   330
         Index           =   44
         Left            =   90
         TabIndex        =   97
         Top             =   5280
         Width           =   690
      End
      Begin VB.ComboBox cmb_Disp2 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   44
         ItemData        =   "FrmTcpServer.frx":32AE
         Left            =   2520
         List            =   "FrmTcpServer.frx":32BB
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   96
         Top             =   4920
         Width           =   615
      End
      Begin VB.ComboBox cmb_Disp1 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   44
         ItemData        =   "FrmTcpServer.frx":32CB
         Left            =   2520
         List            =   "FrmTcpServer.frx":32D8
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   95
         Top             =   4590
         Width           =   615
      End
      Begin VB.TextBox txt_Disp2 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   315
         Index           =   44
         Left            =   90
         TabIndex        =   94
         Text            =   "�����峻 ���� ����"
         Top             =   4920
         Width           =   2430
      End
      Begin VB.TextBox txt_Disp1 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   315
         Index           =   44
         Left            =   90
         TabIndex        =   93
         Text            =   "�ϴ� ����..!!"
         Top             =   4590
         Width           =   2430
      End
      Begin VB.ComboBox cmb_RelayComPort 
         Height          =   330
         Index           =   44
         ItemData        =   "FrmTcpServer.frx":32E8
         Left            =   2295
         List            =   "FrmTcpServer.frx":3307
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   92
         Top             =   3315
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.TextBox txt_RelayPort 
         Height          =   330
         Index           =   44
         Left            =   1635
         TabIndex        =   91
         Text            =   "1100"
         Top             =   3315
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.ComboBox cmb_DispComPort 
         Height          =   330
         Index           =   44
         ItemData        =   "FrmTcpServer.frx":3326
         Left            =   2295
         List            =   "FrmTcpServer.frx":3345
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   90
         Top             =   2925
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.TextBox txt_DeviceIP 
         Height          =   330
         Index           =   44
         Left            =   90
         TabIndex        =   89
         Text            =   "192.168.0.215"
         ToolTipText     =   "�����Ʈ IP�ּҸ� �Է��ϼ���"
         Top             =   2925
         Width           =   1515
      End
      Begin VB.TextBox txt_DispPort 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   330
         Index           =   44
         Left            =   1635
         TabIndex        =   88
         Text            =   "1000"
         Top             =   2925
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.ComboBox cmb_LPRMode 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   330
         Index           =   44
         ItemData        =   "FrmTcpServer.frx":3364
         Left            =   90
         List            =   "FrmTcpServer.frx":3371
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   87
         Top             =   1215
         Visible         =   0   'False
         Width           =   1530
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
         Index           =   44
         Left            =   2265
         TabIndex        =   86
         Top             =   3855
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.TextBox txt_GateName 
         Height          =   315
         Index           =   44
         Left            =   945
         TabIndex        =   85
         Text            =   "����"
         Top             =   600
         Width           =   1725
      End
      Begin VB.TextBox txt_LPRPort 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   330
         Index           =   44
         Left            =   1650
         TabIndex        =   84
         Text            =   "10105"
         Top             =   1440
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.TextBox txt_LPRIP 
         Height          =   330
         Index           =   44
         Left            =   60
         TabIndex        =   83
         Text            =   "192.168.0.204"
         Top             =   1815
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.ComboBox cmb_Inout 
         Height          =   330
         Index           =   44
         ItemData        =   "FrmTcpServer.frx":3388
         Left            =   930
         List            =   "FrmTcpServer.frx":3392
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   82
         Top             =   240
         Width           =   1725
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   8
         X1              =   60
         X2              =   3135
         Y1              =   4380
         Y2              =   4380
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
         BackColor       =   &H00E0E0E0&
         Caption         =   "Relay"
         Height          =   210
         Index           =   18
         Left            =   105
         TabIndex        =   106
         Top             =   3435
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ī�޶󿡼� ����"
         Height          =   210
         Index           =   20
         Left            =   105
         TabIndex        =   105
         Top             =   1215
         Visible         =   0   'False
         Width           =   2190
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "��Ī"
         Height          =   210
         Index           =   21
         Left            =   270
         TabIndex        =   104
         Top             =   645
         Width           =   495
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dispaly"
         Height          =   210
         Index           =   22
         Left            =   105
         TabIndex        =   103
         Top             =   3240
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "����"
         Height          =   285
         Index           =   4
         Left            =   270
         TabIndex        =   102
         Top             =   300
         Width           =   495
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "����� ������"
         Height          =   285
         Index           =   9
         Left            =   120
         TabIndex        =   101
         Top             =   2310
         Width           =   1665
      End
   End
   Begin VB.CheckBox chk_UseYN 
      BackColor       =   &H00E0E0E0&
      Caption         =   "���"
      Height          =   195
      Index           =   44
      Left            =   23070
      TabIndex        =   80
      Top             =   6510
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ī�޶�(LPR) ��Ź��"
      Enabled         =   0   'False
      Height          =   885
      Left            =   20475
      TabIndex        =   78
      Top             =   5520
      Visible         =   0   'False
      Width           =   2145
      Begin VB.ComboBox cmb_LPRMode 
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         ItemData        =   "FrmTcpServer.frx":33A2
         Left            =   180
         List            =   "FrmTcpServer.frx":33AF
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   79
         Top             =   420
         Width           =   1530
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ȭ����"
      ForeColor       =   &H00FF0000&
      Height          =   900
      Left            =   1725
      TabIndex        =   76
      Top             =   1635
      Width           =   1965
      Begin VB.ComboBox Cmb_Window 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "FrmTcpServer.frx":33C6
         Left            =   60
         List            =   "FrmTcpServer.frx":33C8
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   77
         Top             =   390
         Width           =   1845
      End
   End
   Begin VB.TextBox txt_RelayPort 
      Height          =   330
      Index           =   0
      Left            =   945
      TabIndex        =   75
      Text            =   "1100"
      Top             =   13770
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.ComboBox cmb_DispComPort 
      Height          =   330
      Index           =   0
      ItemData        =   "FrmTcpServer.frx":33CA
      Left            =   1605
      List            =   "FrmTcpServer.frx":33E9
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   74
      Top             =   13380
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.TextBox txt_DispPort 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   330
      Index           =   0
      Left            =   945
      TabIndex        =   73
      Text            =   "1000"
      Top             =   13380
      Visible         =   0   'False
      Width           =   630
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
      Left            =   1575
      TabIndex        =   72
      Top             =   14310
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.TextBox txt_LPRPort 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   330
      Index           =   0
      Left            =   1650
      TabIndex        =   71
      Text            =   "10101"
      Top             =   11895
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.TextBox txt_LPRIP 
      Height          =   330
      Index           =   0
      Left            =   60
      TabIndex        =   70
      Text            =   "192.168.0.201"
      Top             =   12270
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.ComboBox cmb_RelayComPort 
      Height          =   330
      Index           =   1
      ItemData        =   "FrmTcpServer.frx":3408
      Left            =   4035
      List            =   "FrmTcpServer.frx":3427
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   69
      Top             =   13860
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.TextBox txt_RelayPort 
      Height          =   330
      Index           =   1
      Left            =   3375
      TabIndex        =   68
      Text            =   "1100"
      Top             =   13860
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.ComboBox cmb_DispComPort 
      Height          =   330
      Index           =   1
      ItemData        =   "FrmTcpServer.frx":3446
      Left            =   4035
      List            =   "FrmTcpServer.frx":3465
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   67
      Top             =   13470
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.TextBox txt_DispPort 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   330
      Index           =   1
      Left            =   3375
      TabIndex        =   66
      Text            =   "1000"
      Top             =   13470
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.ComboBox cmb_LPRMode 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   330
      Index           =   1
      ItemData        =   "FrmTcpServer.frx":3484
      Left            =   2730
      List            =   "FrmTcpServer.frx":3491
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   65
      Top             =   11985
      Visible         =   0   'False
      Width           =   1530
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
      Left            =   3990
      TabIndex        =   64
      Top             =   14370
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.TextBox txt_LPRPort 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   330
      Index           =   1
      Left            =   4290
      TabIndex        =   63
      Text            =   "10102"
      Top             =   11985
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.TextBox txt_LPRIP 
      Height          =   330
      Index           =   1
      Left            =   2715
      TabIndex        =   62
      Text            =   "192.168.0.202"
      Top             =   12360
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.TextBox txt_LPRIP 
      Height          =   330
      Index           =   2
      Left            =   5370
      TabIndex        =   61
      Text            =   "192.168.0.203"
      Top             =   12390
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.TextBox txt_LPRPort 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   330
      Index           =   2
      Left            =   6945
      TabIndex        =   60
      Text            =   "10103"
      Top             =   12015
      Visible         =   0   'False
      Width           =   630
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
      Left            =   6660
      TabIndex        =   59
      Top             =   14385
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.ComboBox cmb_LPRMode 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   330
      Index           =   2
      ItemData        =   "FrmTcpServer.frx":34A8
      Left            =   5385
      List            =   "FrmTcpServer.frx":34B5
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   58
      Top             =   12015
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.TextBox txt_DispPort 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   330
      Index           =   2
      Left            =   6030
      TabIndex        =   57
      Text            =   "1000"
      Top             =   13500
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.ComboBox cmb_DispComPort 
      Height          =   330
      Index           =   2
      ItemData        =   "FrmTcpServer.frx":34CC
      Left            =   6690
      List            =   "FrmTcpServer.frx":34EB
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   56
      Top             =   13500
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.TextBox txt_RelayPort 
      Height          =   330
      Index           =   2
      Left            =   6030
      TabIndex        =   55
      Text            =   "1100"
      Top             =   13890
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.ComboBox cmb_RelayComPort 
      Height          =   330
      Index           =   2
      ItemData        =   "FrmTcpServer.frx":350A
      Left            =   6690
      List            =   "FrmTcpServer.frx":3529
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   54
      Top             =   13890
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.TextBox txt_LPRIP 
      Height          =   330
      Index           =   3
      Left            =   7830
      TabIndex        =   53
      Text            =   "192.168.0.204"
      Top             =   12375
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.TextBox txt_LPRPort 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   330
      Index           =   3
      Left            =   9405
      TabIndex        =   52
      Text            =   "10104"
      Top             =   12000
      Visible         =   0   'False
      Width           =   630
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
      Left            =   9180
      TabIndex        =   51
      Top             =   14370
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.ComboBox cmb_LPRMode 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   330
      Index           =   3
      ItemData        =   "FrmTcpServer.frx":3548
      Left            =   7845
      List            =   "FrmTcpServer.frx":3555
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   50
      Top             =   12000
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.TextBox txt_DispPort 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   330
      Index           =   3
      Left            =   8550
      TabIndex        =   49
      Text            =   "1000"
      Top             =   13485
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.ComboBox cmb_DispComPort 
      Height          =   330
      Index           =   3
      ItemData        =   "FrmTcpServer.frx":356C
      Left            =   9210
      List            =   "FrmTcpServer.frx":358B
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   48
      Top             =   13485
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.TextBox txt_RelayPort 
      Height          =   330
      Index           =   3
      Left            =   8550
      TabIndex        =   47
      Text            =   "1100"
      Top             =   13875
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.ComboBox cmb_RelayComPort 
      Height          =   330
      Index           =   3
      ItemData        =   "FrmTcpServer.frx":35AA
      Left            =   9210
      List            =   "FrmTcpServer.frx":35C9
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   46
      Top             =   13875
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.ComboBox cmb_RelayComPort 
      Height          =   330
      Index           =   0
      ItemData        =   "FrmTcpServer.frx":35E8
      Left            =   1620
      List            =   "FrmTcpServer.frx":3607
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   45
      Top             =   12315
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Caption         =   "�ΰ� ���"
      Height          =   840
      Left            =   45
      TabIndex        =   37
      Top             =   390
      Width           =   19260
      Begin VB.CommandButton Command10 
         Caption         =   "����Ͼ˸�"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3315
         TabIndex        =   362
         Top             =   285
         Width           =   1260
      End
      Begin VB.TextBox txt_Vendor 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7905
         TabIndex        =   341
         Text            =   "���ڸ���"
         ToolTipText     =   "��ü���� �Է��ϼ���"
         Top             =   300
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.CommandButton Command2 
         Caption         =   "������"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   10815
         TabIndex        =   339
         Top             =   300
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.TextBox txt_SiteName 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   9525
         TabIndex        =   337
         ToolTipText     =   "������� �Է��ϼ���"
         Top             =   300
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.CommandButton Command6 
         Caption         =   "���μ���"
         BeginProperty Font 
            Name            =   "���� ���"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   165
         TabIndex        =   41
         Top             =   285
         Width           =   1260
      End
      Begin VB.CommandButton Command1 
         Caption         =   "����/�����ǹ���"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1515
         TabIndex        =   40
         ToolTipText     =   "�����ġ ""������(Ǯ�÷�)_FW7"" �� ������ ���λ��� ���� �����մϴ�"
         Top             =   285
         Width           =   1710
      End
      Begin VB.CommandButton cmd_Certify 
         Caption         =   "�����ʿ�"
         Height          =   435
         Left            =   11955
         TabIndex        =   39
         ToolTipText     =   "����Ű���� ��� ������������"
         Top             =   285
         Width           =   1065
      End
      Begin VB.TextBox txt_CertifyKey 
         Height          =   435
         IMEMode         =   3  '��� ����
         Left            =   13035
         PasswordChar    =   "*"
         TabIndex        =   38
         Text            =   "admin0808"
         ToolTipText     =   "����Ű �Է��ϼ���"
         Top             =   285
         Width           =   765
      End
      Begin Threed.SSCommand cmd_Button 
         Height          =   555
         Index           =   0
         Left            =   16455
         TabIndex        =   42
         Top             =   180
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   979
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
         Picture         =   "FrmTcpServer.frx":3626
      End
      Begin Threed.SSCommand cmd_Button 
         Height          =   555
         Index           =   1
         Left            =   17835
         TabIndex        =   43
         Top             =   180
         Width           =   1335
         _Version        =   65536
         _ExtentX        =   2355
         _ExtentY        =   979
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
         Picture         =   "FrmTcpServer.frx":3977
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         Caption         =   "��ü��"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7335
         TabIndex        =   342
         Top             =   420
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�����"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8955
         TabIndex        =   338
         Top             =   390
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label lbl_CertifyLimitDate 
         BackColor       =   &H000000FF&
         Caption         =   "����Ⱓ:2019-01-01"
         BeginProperty Font 
            Name            =   "���� ���"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   13860
         TabIndex        =   44
         ToolTipText     =   "����Ⱓ �̳� ������������. ����Ⱓ ���Ŀ��� ���ܱⰡ �����۵����� �ʽ��ϴ�."
         Top             =   315
         Width           =   2370
      End
   End
   Begin VB.CommandButton cmd_Svr 
      Caption         =   "SET"
      Height          =   360
      Left            =   16605
      TabIndex        =   36
      Top             =   13320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtPort 
      BackColor       =   &H00C0C0C0&
      Height          =   315
      Left            =   15945
      TabIndex        =   35
      Text            =   "10100"
      Top             =   13395
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton CmdSvr 
      Caption         =   "SET"
      Height          =   360
      Index           =   0
      Left            =   16845
      TabIndex        =   34
      Top             =   13980
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton CmdSvr 
      Caption         =   "SET"
      Height          =   360
      Index           =   1
      Left            =   18120
      TabIndex        =   33
      Top             =   14085
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      Caption         =   "�ⱸ��������"
      ForeColor       =   &H00FF0000&
      Height          =   900
      Left            =   17505
      TabIndex        =   27
      Top             =   1635
      Width           =   1800
      Begin VB.TextBox TxtAspIp 
         Height          =   315
         Left            =   180
         TabIndex        =   28
         Text            =   "255.255.255.255"
         ToolTipText     =   "�ⱸ��������� ip�ּҸ� ��������"
         Top             =   495
         Width           =   1395
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         Caption         =   "������"
         Height          =   255
         Left            =   225
         TabIndex        =   29
         Top             =   285
         Width           =   885
      End
   End
   Begin VB.ComboBox cmb_RelayComPort 
      Height          =   330
      Index           =   4
      ItemData        =   "FrmTcpServer.frx":3CC8
      Left            =   11925
      List            =   "FrmTcpServer.frx":3CE7
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   25
      Top             =   13815
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.TextBox txt_RelayPort 
      Height          =   330
      Index           =   4
      Left            =   11265
      TabIndex        =   24
      Text            =   "1100"
      Top             =   13815
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.ComboBox cmb_DispComPort 
      Height          =   330
      Index           =   4
      ItemData        =   "FrmTcpServer.frx":3D06
      Left            =   11925
      List            =   "FrmTcpServer.frx":3D25
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   23
      Top             =   13425
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.TextBox txt_DispPort 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   330
      Index           =   4
      Left            =   11265
      TabIndex        =   22
      Text            =   "1000"
      Top             =   13425
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.ComboBox cmb_LPRMode 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   330
      Index           =   4
      ItemData        =   "FrmTcpServer.frx":3D44
      Left            =   10560
      List            =   "FrmTcpServer.frx":3D51
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   21
      Top             =   11940
      Visible         =   0   'False
      Width           =   1530
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
      Left            =   11895
      TabIndex        =   20
      Top             =   14310
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.TextBox txt_LPRPort 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   330
      Index           =   4
      Left            =   12120
      TabIndex        =   19
      Text            =   "10105"
      Top             =   11940
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.TextBox txt_LPRIP 
      Height          =   330
      Index           =   4
      Left            =   10545
      TabIndex        =   18
      Text            =   "192.168.0.204"
      Top             =   12315
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.ComboBox cmb_RelayComPort 
      Height          =   330
      Index           =   5
      ItemData        =   "FrmTcpServer.frx":3D68
      Left            =   14385
      List            =   "FrmTcpServer.frx":3D87
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   17
      Top             =   13830
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.TextBox txt_RelayPort 
      Height          =   330
      Index           =   5
      Left            =   13725
      TabIndex        =   16
      Text            =   "1100"
      Top             =   13830
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.ComboBox cmb_DispComPort 
      Height          =   330
      Index           =   5
      ItemData        =   "FrmTcpServer.frx":3DA6
      Left            =   14385
      List            =   "FrmTcpServer.frx":3DC5
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   15
      Top             =   13440
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.TextBox txt_DispPort 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   330
      Index           =   5
      Left            =   13725
      TabIndex        =   14
      Text            =   "1000"
      Top             =   13440
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.ComboBox cmb_LPRMode 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   330
      Index           =   5
      ItemData        =   "FrmTcpServer.frx":3DE4
      Left            =   13020
      List            =   "FrmTcpServer.frx":3DF1
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   13
      Top             =   11955
      Visible         =   0   'False
      Width           =   1530
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
      Index           =   5
      Left            =   14355
      TabIndex        =   12
      Top             =   14325
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.TextBox txt_LPRPort 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   330
      Index           =   5
      Left            =   14580
      TabIndex        =   11
      Text            =   "10106"
      Top             =   11955
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.TextBox txt_LPRIP 
      Height          =   330
      Index           =   5
      Left            =   13005
      TabIndex        =   10
      Text            =   "192.168.0.204"
      Top             =   12330
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Timer DB_Connect_Timer 
      Enabled         =   0   'False
      Left            =   21195
      Top             =   -60
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00E0E0E0&
      Caption         =   "�����ġ ����"
      ForeColor       =   &H00000000&
      Height          =   885
      Left            =   6330
      TabIndex        =   8
      Top             =   1635
      Width           =   5325
      Begin VB.ComboBox cmb_DispToggleTime 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         ItemData        =   "FrmTcpServer.frx":3E08
         Left            =   3615
         List            =   "FrmTcpServer.frx":3E27
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   354
         ToolTipText     =   "������ �̵� �ӵ�����(���ڰ� ���� ���� ���� �̵��մϴ�), ���� �� �Ϲݹ�ư ��������"
         Top             =   390
         Width           =   870
      End
      Begin VB.ComboBox cmb_DispToggleCount 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         ItemData        =   "FrmTcpServer.frx":3E4F
         Left            =   4500
         List            =   "FrmTcpServer.frx":3E6E
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   353
         ToolTipText     =   "������ �̵� �ӵ�����(���ڰ� ���� ���� ���� �̵��մϴ�), ���� �� �Ϲݹ�ư ��������"
         Top             =   390
         Width           =   750
      End
      Begin VB.ComboBox cmb_DisplayMode 
         Height          =   330
         Index           =   0
         ItemData        =   "FrmTcpServer.frx":3E96
         Left            =   2025
         List            =   "FrmTcpServer.frx":3EA0
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   346
         Top             =   390
         Width           =   705
      End
      Begin VB.ComboBox Cmb_Display_Direct 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "FrmTcpServer.frx":3EAE
         Left            =   2745
         List            =   "FrmTcpServer.frx":3EB0
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   324
         Top             =   390
         Width           =   855
      End
      Begin VB.ComboBox Cmb_Display 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "FrmTcpServer.frx":3EB2
         Left            =   120
         List            =   "FrmTcpServer.frx":3EB4
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   9
         Top             =   390
         Width           =   1890
      End
      Begin VB.Label lbl_DispToggleTime 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�����ð�(s)"
         Height          =   255
         Left            =   3600
         TabIndex        =   356
         Top             =   195
         Width           =   870
      End
      Begin VB.Label lbl_DispToggleCount 
         BackColor       =   &H00E0E0E0&
         Caption         =   "���Ƚ��"
         Height          =   255
         Left            =   4545
         TabIndex        =   355
         Top             =   195
         Visible         =   0   'False
         Width           =   750
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00E0E0E0&
      Caption         =   "������������"
      ForeColor       =   &H00FF0000&
      Height          =   900
      Left            =   15180
      TabIndex        =   3
      Top             =   1635
      Width           =   2280
      Begin VB.TextBox TxtGraceTime 
         Height          =   315
         Left            =   135
         TabIndex        =   5
         Text            =   "10"
         ToolTipText     =   "������ ���������� ���ð�"
         Top             =   510
         Width           =   705
      End
      Begin VB.TextBox TxtReturnTime 
         Height          =   315
         Left            =   1245
         TabIndex        =   4
         Text            =   "5"
         ToolTipText     =   "ȸ���ð�"
         Top             =   510
         Width           =   705
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         Caption         =   "�׷��̽�Ÿ��"
         Height          =   255
         Left            =   105
         TabIndex        =   7
         Top             =   270
         Width           =   1125
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ȸ���ð�(��)"
         Height          =   255
         Left            =   1215
         TabIndex        =   6
         Top             =   270
         Width           =   1035
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00E0E0E0&
      Caption         =   "����̽� ����"
      Height          =   885
      Left            =   3750
      TabIndex        =   0
      Top             =   1635
      Width           =   2535
      Begin VB.ComboBox cmb_DeviceMode 
         Height          =   330
         Index           =   0
         ItemData        =   "FrmTcpServer.frx":3EB6
         Left            =   1470
         List            =   "FrmTcpServer.frx":3EC0
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   345
         Top             =   420
         Width           =   1005
      End
      Begin VB.ComboBox cmb_Board 
         Height          =   330
         ItemData        =   "FrmTcpServer.frx":3ECE
         Left            =   60
         List            =   "FrmTcpServer.frx":3ED8
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   1
         Top             =   420
         Width           =   1380
      End
   End
   Begin VB.Timer Timer_Certify 
      Interval        =   5000
      Left            =   13545
      Top             =   810
   End
   Begin MSCommLib.MSComm MSComm 
      Index           =   0
      Left            =   20295
      Top             =   4740
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   495
      Left            =   30
      TabIndex        =   30
      Top             =   8130
      Width           =   19260
      _Version        =   65536
      _ExtentX        =   33972
      _ExtentY        =   873
      _StockProps     =   15
      Caption         =   "�ý��� �α�â"
      BackColor       =   32896
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton Command8 
         Caption         =   "Ŭ����"
         Height          =   315
         Left            =   17250
         TabIndex        =   32
         Top             =   105
         Width           =   975
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00008080&
         Caption         =   "Refresh"
         Height          =   225
         Left            =   18330
         TabIndex        =   31
         Top             =   150
         Value           =   1  'Ȯ��
         Width           =   945
      End
   End
   Begin MSWinsockLib.Winsock HomeSock 
      Left            =   20730
      Top             =   2625
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock LPR1_sock 
      Index           =   3
      Left            =   21570
      Top             =   945
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock LPR1_sock 
      Index           =   2
      Left            =   21150
      Top             =   945
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock LPR1_sock 
      Index           =   1
      Left            =   20730
      Top             =   945
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Gate1_sock 
      Index           =   3
      Left            =   21570
      Top             =   1785
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Gate1_sock 
      Index           =   2
      Left            =   21150
      Top             =   1785
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Gate1_sock 
      Index           =   1
      Left            =   20730
      Top             =   1785
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Gate1_sock 
      Index           =   0
      Left            =   20310
      Top             =   1785
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Disp1_sock 
      Index           =   0
      Left            =   20310
      Top             =   1365
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock LPR1_sock 
      Index           =   0
      Left            =   20310
      Top             =   945
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Disp1_sock 
      Index           =   1
      Left            =   20730
      Top             =   1365
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Disp1_sock 
      Index           =   2
      Left            =   21150
      Top             =   1365
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Disp1_sock 
      Index           =   3
      Left            =   21570
      Top             =   1365
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock RemoteR_sock 
      Left            =   21990
      Top             =   2625
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock RemoteS_sock 
      Left            =   21570
      Top             =   2625
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock MvrSock 
      Left            =   20310
      Top             =   2625
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock LPR1_sock 
      Index           =   4
      Left            =   21990
      Top             =   945
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Disp1_sock 
      Index           =   4
      Left            =   21990
      Top             =   1365
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Gate1_sock 
      Index           =   4
      Left            =   21990
      Top             =   1785
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Gate1_sock 
      Index           =   5
      Left            =   22410
      Top             =   1785
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Disp1_sock 
      Index           =   5
      Left            =   22410
      Top             =   1365
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock LPR1_sock 
      Index           =   5
      Left            =   22410
      Top             =   945
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock LPR_Send_sock 
      Index           =   0
      Left            =   20310
      Top             =   525
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock LPR_Send_sock 
      Index           =   1
      Left            =   20730
      Top             =   525
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock LPR_Send_sock 
      Index           =   2
      Left            =   21150
      Top             =   525
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock LPR_Send_sock 
      Index           =   3
      Left            =   21570
      Top             =   525
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock LPR_Send_sock 
      Index           =   4
      Left            =   21990
      Top             =   525
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock LPR_Send_sock 
      Index           =   5
      Left            =   22410
      Top             =   525
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSCommLib.MSComm MSComm 
      Index           =   1
      Left            =   20865
      Top             =   4740
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm 
      Index           =   2
      Left            =   21435
      Top             =   4740
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm 
      Index           =   3
      Left            =   22005
      Top             =   4740
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm 
      Index           =   4
      Left            =   22575
      Top             =   4740
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm 
      Index           =   5
      Left            =   23145
      Top             =   4740
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSWinsockLib.Winsock DeviceR_sock 
      Left            =   22410
      Top             =   2625
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock Reset_sock 
      Index           =   0
      Left            =   20310
      Top             =   2205
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Reset_sock 
      Index           =   1
      Left            =   20730
      Top             =   2205
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Reset_sock 
      Index           =   2
      Left            =   21150
      Top             =   2205
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Reset_sock 
      Index           =   3
      Left            =   21570
      Top             =   2205
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Reset_sock 
      Index           =   4
      Left            =   21990
      Top             =   2205
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Reset_sock 
      Index           =   5
      Left            =   22410
      Top             =   2205
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin LPR_PARKING_HOST.Server Server_GateAgentR 
      Index           =   1
      Left            =   20730
      Top             =   3150
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin LPR_PARKING_HOST.Server Server_GateAgentR 
      Index           =   2
      Left            =   21150
      Top             =   3150
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin LPR_PARKING_HOST.Server Server_GateAgentR 
      Index           =   3
      Left            =   21570
      Top             =   3150
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin LPR_PARKING_HOST.Server Server_GateAgentR 
      Index           =   4
      Left            =   21990
      Top             =   3150
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin LPR_PARKING_HOST.Server Server_GateAgentR 
      Index           =   5
      Left            =   22410
      Top             =   3150
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin MSWinsockLib.Winsock Winsock_GateAgentR 
      Index           =   1
      Left            =   20730
      Top             =   3585
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock Winsock_GateAgentR 
      Index           =   2
      Left            =   21150
      Top             =   3585
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock Winsock_GateAgentR 
      Index           =   3
      Left            =   21570
      Top             =   3585
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock Winsock_GateAgentR 
      Index           =   4
      Left            =   21990
      Top             =   3585
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock Winsock_GateAgentR 
      Index           =   5
      Left            =   22410
      Top             =   3585
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   5265
      Left            =   23970
      TabIndex        =   360
      Top             =   525
      Visible         =   0   'False
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   9287
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   16777215
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
      NumItems        =   0
   End
   Begin VB.Label lbl_SiteName 
      BackColor       =   &H0080C0FF&
      Caption         =   "lbl_SiteName"
      Height          =   225
      Left            =   20310
      TabIndex        =   344
      Top             =   4470
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label lbl_Vendor 
      BackColor       =   &H0080C0FF&
      Caption         =   "lbl_Vendor"
      Height          =   225
      Left            =   20310
      TabIndex        =   343
      Top             =   4200
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404040&
      BackStyle       =   0  '����
      Caption         =   " # ���� #  �ý��� ������ �ܿ� ���� ���� ����..!!"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   75
      TabIndex        =   165
      Top             =   75
      Width           =   5145
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Relay"
      Height          =   210
      Index           =   1
      Left            =   105
      TabIndex        =   164
      Top             =   13890
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dispaly"
      Height          =   210
      Index           =   0
      Left            =   105
      TabIndex        =   163
      Top             =   13695
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Relay"
      Height          =   210
      Index           =   3
      Left            =   2745
      TabIndex        =   162
      Top             =   13980
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ī�޶󿡼� ����"
      Height          =   210
      Index           =   5
      Left            =   2745
      TabIndex        =   161
      Top             =   11760
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dispaly"
      Height          =   210
      Index           =   7
      Left            =   2745
      TabIndex        =   160
      Top             =   13785
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dispaly"
      Height          =   210
      Index           =   12
      Left            =   5400
      TabIndex        =   159
      Top             =   13815
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ī�޶󿡼� ����"
      Height          =   210
      Index           =   10
      Left            =   5400
      TabIndex        =   158
      Top             =   11790
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Relay"
      Height          =   210
      Index           =   8
      Left            =   5400
      TabIndex        =   157
      Top             =   14010
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dispaly"
      Height          =   210
      Index           =   17
      Left            =   7860
      TabIndex        =   156
      Top             =   13800
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ī�޶󿡼� ����"
      Height          =   210
      Index           =   15
      Left            =   7860
      TabIndex        =   155
      Top             =   11775
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Relay"
      Height          =   210
      Index           =   13
      Left            =   7860
      TabIndex        =   154
      Top             =   13995
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "��Ʈ"
      Height          =   255
      Left            =   15975
      TabIndex        =   153
      Top             =   13170
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Relay"
      Height          =   210
      Index           =   27
      Left            =   10575
      TabIndex        =   152
      Top             =   13935
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ī�޶󿡼� ����"
      Height          =   210
      Index           =   28
      Left            =   10575
      TabIndex        =   151
      Top             =   11715
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dispaly"
      Height          =   210
      Index           =   29
      Left            =   10575
      TabIndex        =   150
      Top             =   13740
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Relay"
      Height          =   210
      Index           =   30
      Left            =   13035
      TabIndex        =   149
      Top             =   13950
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ī�޶󿡼� ����"
      Height          =   210
      Index           =   31
      Left            =   13035
      TabIndex        =   148
      Top             =   11730
      Visible         =   0   'False
      Width           =   2160
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dispaly"
      Height          =   210
      Index           =   32
      Left            =   13035
      TabIndex        =   147
      Top             =   13755
      Visible         =   0   'False
      Width           =   585
   End
End
Attribute VB_Name = "FrmTcpServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const white = &H80000005
Const grey = &H8000000F

Public Enum typeGateKind
    InGate = 0
    OutGate = 1
End Enum
Public eGateKind As typeGateKind
Public frontno As Integer


Private Sub ApsS_sock_Connect()

Dim sdata As String
Dim bData() As Byte
Dim i As Integer

On Error GoTo Err_p

    sdata = Glo_APS_Str
    ReDim bData(Len(sdata) - 1) As Byte
    bData = StrConv(sdata, vbFromUnicode)
    ApsS_sock.SendData bData
    Glo_APS_Str = ""

Exit Sub

Err_p:
    Call DataLogger(" [ApsS_sock_Connect ���ν���] �������� : " & Err.Description)

End Sub

Private Sub Aps_UDP_DataArrival(ByVal bytesTotal As Long)
    Dim sdata As String
    Dim Tmp_Path As String
    Dim i, gateNo As Integer
    Dim carnum As String
    Dim cmd As String
    
    Dim qry As String
    Dim rs As ADODB.Recordset
    Dim bQryResult As Boolean
    
    
On Error GoTo Err_p
    
    
    If (bytesTotal > 500) Then
        'DebugLogger ("Aps ������ �ʰ�����(������) : " & bytesTotal)
        Exit Sub
    End If
    
    
    Aps_UDP.GetData sdata, , bytesTotal
    Call DataLogger("Aps_UDP UDP Port " & FormatNumber(bytesTotal, 0, , , vbTrue) & " bytes recieved." & "    " & sdata)
    
    
    
    
With FrmAccnt
    
    
    
cmd = Mid(sdata, 1, 2)
    
    
    
    Select Case cmd
            Case CM_NOPAY
                 .LblMsg.Caption = "����ó��"
                 .LblAps(9).Caption = MidH(sdata, 3, LenH(sdata) - 2)
            Case CM_START ' ��������� START
                 .LblMsg.Caption = "��������� START"
            Case CM_END ' ��������� END
                 .LblMsg.Caption = ""
                 .LblAps(8).Visible = True
                 .LblAps(9).Visible = True
                 .LblAps(8).Caption = "�����..."
                 
                For i = 0 To 7
                    .LblAps(i).Visible = False
                    .LblAps(i).Caption = ""
                Next i
                .Image1.Picture = LoadPicture(App.Path & "\Image\asp_small1.bmp")
                .Image2.Picture = LoadPicture(App.Path & "\NoCar.jpg")
                APS_INFO_CarNo = ""
                 
            Case CM_RESPONSE ' ȣ��Ʈ ��� ����
                 .LblMsg.Caption = MidH(sdata, 3, LenH(sdata) - 2)
            Case CM_JUNGSANCANCEL ' ������ҹ�ư ����
                 .LblMsg.Caption = "������ҹ�ư ����"
            Case CM_CHANGEOUTERR ' �Ž����� ���⿡��
            Case CM_DISPENSER1000ERR ' 1000���� �������⿡��
            
            Case CM_DISPENSER5000ERR ' 5000���� �������⿡��
            Case CM_COINERR ' ���α⿡��
            Case CM_BILLERR ' �����νı⿡��
            Case CM_CAROUT ' ���� ����
                     For i = 0 To 7
                         .LblAps(i).Visible = True
                     Next i
                     .LblMsg.Caption = ""
                     
                    If (APS_INFO_CarNo = "") Then
                        'ù��° ���Žÿ��� �̹��� �ε�
                        qry = "Select * From tb_now Where CAR_NO = '" & Trim(Mid(sdata, 3, 12)) & "' Order By PASS_DATE Desc"
                        Set rs = New ADODB.Recordset
'                        rs.Open Qry, adoConn
                        bQryResult = DataBaseQuery(rs, adoConn, qry, False)
                        If (bQryResult = False) Then
                            ListData.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    ��Ʈ��ũ �� DB ���˹ٶ��ϴ�", 0
                            Call DataLogger("[FrmReg]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
                            Exit Sub
                        End If
                        
                        If Not (rs.EOF) Then
                            If (IsFile(rs!pass_image) = True) Then
                                .Image2.Picture = LoadPicture(rs!pass_image)
                            Else
                                .Image2.Picture = LoadPicture(App.Path & "\NoCar.jpg")
                            End If
                        End If
                    End If
                        Dim iTime As Long
                        Dim sHour As String
                        Dim sMin As String
                        Dim sTime As String
                     .LblAps(8).Caption = ""
                     .LblAps(9).Caption = ""
                     .LblAps(8).Visible = False
                     .LblAps(9).Visible = False
                     .Image1.Picture = LoadPicture(App.Path & "\Image\asp_small2.bmp")
                     .LblAps(7).Caption = Trim(Mid(sdata, 3, 12))
                     .LblAps(0).Caption = "�����Ͻ� : " & Trim(Mid(sdata, 15, 16))
                     .LblAps(1).Caption = "�����Ͻ� : " & Trim(Mid(sdata, 31, 16))
                     '.LblAps(2).Caption = "�����ð� : " & Trim(Mid(sdata, 47, 6))
                     iTime = Val(Trim(Mid(sdata, 47, 6)))
                        If (Int(iTime / 60) > 0) Then
                            sHour = CStr(Int(iTime / 60)) & "�ð� "
                        End If
                        If (iTime Mod 60 > 0) Then
                            sMin = CStr(iTime Mod 60) & "�� "
                        End If
                     sTime = sHour & sMin
                     .LblAps(2).Caption = "�����ð� : " & sTime
                     
                     .LblAps(3).Caption = "��   �� : " & Space(10 - Len(Trim(Mid(sdata, 53, 6)))) & Trim(Mid(sdata, 53, 6))
                     .LblAps(4).Caption = "��   �� : " & Space(10 - Len(Trim(Mid(sdata, 59, 6)))) & Trim(Mid(sdata, 59, 6))
                     .LblAps(5).Caption = "��   �� : " & Space(10 - Len(Trim(Mid(sdata, 65, 6)))) & Trim(Mid(sdata, 65, 6))
                     
                     If (Val(Mid(sdata, 71, 6)) <= 0) Then
                        '�Ž����� �߻��� !!!
                        .LblAps(6).Caption = "��   �� : " & "         0"
                     Else
                        .LblAps(6).Caption = "��   �� : " & Space(10 - Len(Trim(Mid(sdata, 71, 6)))) & Trim(Mid(sdata, 71, 6))
                     End If
                     
                     If (Val(Mid(sdata, 71, 6)) <= 0) Then
                         .LblMsg.Caption = "����Ϸ� ������ �߱޴����..."
                         Call LISTBOX_PutString(.List_SALE, " " & .LblAps(7).Caption & ", �����ð�:" & sTime & ", ���:" & Trim(Mid(sdata, 59, 6)) & ", ����:" & Trim(Mid(sdata, 53, 6)) & ", ����:" & Trim(Mid(sdata, 65, 6)) & Right(sdata, Len(sdata) - 76))
                     End If
                     Call Read_Account
                     
                     
                     '�������� ����Ʈ�ڽ� ��� ����
                    If (APS_INFO_CarNo = "") Then
                        APS_INFO_CarNo = .LblAps(7).Caption
                        Call LISTBOX_PutString(.List_OP, " ---------------------------------------------------------------------")
                        Call LISTBOX_PutString(.List_OP, " " & .LblAps(7).Caption & ", ���:" & Trim(Mid(sdata, 59, 6)) & " : �������")
                    Else
                        Call LISTBOX_PutString(.List_OP, " " & .LblAps(7).Caption & ", ���:" & Trim(Mid(sdata, 59, 6)) & ", ����:" & Trim(Mid(sdata, 53, 6)) & ", ����:" & Trim(Mid(sdata, 65, 6)) & ", �ܾ�:" & Trim(Mid(sdata, 71, 6)) & " " & Right(sdata, Len(sdata) - 76))
                    End If
                    '�������� ����Ʈ�ڽ� ��� ��
                    
            Case CM_FILTER ' ���� ������ ���͸��� ���� ���ݵǾ���
            Case CM_NOCAR ' ���� ������ ����
                 .LblMsg.Caption = "���� ���� ����"
                 .LblAps(9).Caption = MidH(sdata, 3, LenH(sdata) - 2)
            Case CM_SERVICECARDERR ' ���αǿ���
                 .LblMsg.Caption = MidH(sdata, 3, LenH(sdata) - 2)
            Case CM_CREDITCARDCANCEL ' �ſ�ī�� �������
                 .LblMsg.Caption = MidH(sdata, 3, LenH(sdata) - 2)
    End Select
End With
Exit Sub

Err_p:
    Call DataLogger("[Aps_UDP DataArrival] " & Err.Description)
End Sub

Private Sub chk_ApsYN_Click()
If (chk_ApsYN.value = 0) Then
    Frame7.Enabled = False
    TxtAspIp.BackColor = &HE0E0E0
Else
    Frame7.Enabled = True
    TxtAspIp.BackColor = &H80000005
End If
End Sub






Private Sub chk_GuestYN_Click(Index As Integer)
    If chk_GuestYN(Index).value = 1 Then
        
        cmb_PrintModel(Index).Enabled = True
        cmb_PrintPort(Index).Enabled = True
        cmb_PrintModel(Index).BackColor = &H80000005 'White
        cmb_PrintPort(Index).BackColor = &H80000005
    Else
        cmb_PrintModel(Index).Enabled = False
        cmb_PrintPort(Index).Enabled = False
        cmb_PrintModel(Index).BackColor = &HE0E0E0 'Gray
        cmb_PrintPort(Index).BackColor = &HE0E0E0
    End If
End Sub

Private Sub chk_PreApsYN_Click()
    If (chk_PreApsYN.value = 0) Then
        Frame9.Enabled = False
        TxtGraceTime.BackColor = &HE0E0E0
        TxtReturnTime.BackColor = &HE0E0E0
    Else
        Frame9.Enabled = True
        TxtGraceTime.BackColor = &H80000005
        TxtReturnTime.BackColor = &H80000005
    End If
End Sub



Private Sub cmb_Board_Click()
    Dim i As Integer
    
    If cmb_Board.text = "�����" Or cmb_Board.text = "�ڵ��̳�" Then
            'Frame4.Caption = "����̽� ��Ź��"
            Frame8.Caption = "�����ġ ����"
            
            '�����ġ
            Cmb_Display.Clear
            Cmb_Display.AddItem "������"
            'Cmb_Display.AddItem "������(Full Color)"
            Cmb_Display.AddItem "FND"
            Cmb_Display.AddItem "������(Ǯ�÷�)"
            Cmb_Display.AddItem "������(Ǯ�÷�)_FW7"
            
            If (Glo_Display = "������") Then
                Cmb_Display.ListIndex = 0
            ElseIf (Glo_Display = "FND") Then
                Cmb_Display.ListIndex = 1
            ElseIf (Glo_Display = "������(Ǯ�÷�)") Then
                Cmb_Display.ListIndex = 2
            ElseIf (Glo_Display = "������(Ǯ�÷�)_FW7") Then
                Cmb_Display.ListIndex = 3
            Else
                Cmb_Display.ListIndex = 3
            End If
        
        For i = 0 To MAX_LANE_COUNT - 1
            lbl_DeviceIP(i).Visible = True
            txt_DeviceIP(i).Visible = True
            
            txt_Disp1(i).width = 2430
            txt_Disp2(i).width = 2430
            cmb_Disp1(i).Left = 2520
            cmb_Disp2(i).Left = 2520
            
            If (cmb_Board.text = "�ڵ��̳�") Then
                cmd_DeviceReset(i).Visible = True
            Else
                cmd_DeviceReset(i).Visible = False
            End If
            
            Glo_LPRBoard = cmb_Board.text
            Call Put_Ini("System Config", "LPRBoard", cmb_Board.text)
        Next i
    End If
    
    
    If (cmb_Board.text = "�����") Then
        For i = 0 To MAX_LANE_COUNT - 1
            cmd_DeviceReset(i).Enabled = False
            cmd_DeviceReset(i).Visible = False
            cmd_DispReset(i).Enabled = False
            cmd_DispReset(i).Visible = False
            
            cmd_GateTestDown(i).Visible = False
            cmd_GateTestDown(i).Enabled = False
        Next i

    ElseIf cmb_Board.text = "�ڵ��̳�" Then
        For i = 0 To MAX_LANE_COUNT - 1
            cmd_DeviceReset(i).Enabled = True
            cmd_DeviceReset(i).Visible = True
            cmd_DispReset(i).Enabled = True
            cmd_DispReset(i).Visible = True
            
            cmd_GateTestDown(i).Visible = True
            cmd_GateTestDown(i).Enabled = True
        Next i
    End If
    
End Sub


Private Sub Cmb_Display_Click()
    Dim i As Integer
    
    If (Cmb_Display.text = "������(Ǯ�÷�)_FW7") Then
        For i = 0 To MAX_LANE_COUNT - 1
            cmb_Disp1(i).Clear
            cmb_Disp1(i).AddItem "��"
            cmb_Disp1(i).AddItem "��"
            cmb_Disp1(i).AddItem "Ȳ"
            cmb_Disp1(i).AddItem "��"
            cmb_Disp1(i).AddItem "��"
            cmb_Disp1(i).AddItem "��"
            cmb_Disp1(i).AddItem "��"
            cmb_Disp2(i).Clear
            cmb_Disp2(i).AddItem "��"
            cmb_Disp2(i).AddItem "��"
            cmb_Disp2(i).AddItem "Ȳ"
            cmb_Disp2(i).AddItem "��"
            cmb_Disp2(i).AddItem "��"
            cmb_Disp2(i).AddItem "��"
            cmb_Disp2(i).AddItem "��"
        Next i
    ElseIf (Cmb_Display.text = "������(Ǯ�÷�)") Then
        For i = 0 To MAX_LANE_COUNT - 1
            cmb_Disp1(i).Clear
            cmb_Disp1(i).AddItem "��"
            cmb_Disp1(i).AddItem "��"
            cmb_Disp1(i).AddItem "Ȳ"
            cmb_Disp2(i).Clear
            cmb_Disp2(i).AddItem "��"
            cmb_Disp2(i).AddItem "��"
            cmb_Disp2(i).AddItem "Ȳ"
        Next i
    Else
        For i = 0 To MAX_LANE_COUNT - 1
            cmb_Disp1(i).Clear
            cmb_Disp1(i).AddItem "��"
            cmb_Disp1(i).AddItem "��"
            cmb_Disp1(i).AddItem "Ȳ"
            cmb_Disp2(i).Clear
            cmb_Disp2(i).AddItem "��"
            cmb_Disp2(i).AddItem "��"
            cmb_Disp2(i).AddItem "Ȳ"
        Next i
    End If
    
    cmb_Disp1(0).ListIndex = LANE1_Disp1Color
    cmb_Disp2(0).ListIndex = LANE1_Disp2Color
    cmb_Disp1(1).ListIndex = LANE2_Disp1Color
    cmb_Disp2(1).ListIndex = LANE2_Disp2Color
    cmb_Disp1(2).ListIndex = LANE3_Disp1Color
    cmb_Disp2(2).ListIndex = LANE3_Disp2Color
    cmb_Disp1(3).ListIndex = LANE4_Disp1Color
    cmb_Disp2(3).ListIndex = LANE4_Disp2Color
    cmb_Disp1(4).ListIndex = LANE5_Disp1Color
    cmb_Disp2(4).ListIndex = LANE5_Disp2Color
    cmb_Disp1(5).ListIndex = LANE6_Disp1Color
    cmb_Disp2(5).ListIndex = LANE6_Disp2Color
    
    If (Cmb_Display.text = "������(Ǯ�÷�)_FW7") Then
        Cmb_Display_Direct.Visible = True
        'cmb_DisplayMode(0).Visible = True
        cmb_DispToggleTime.Visible = True
        cmb_DispToggleCount.Visible = True
        lbl_DispToggleTime.Visible = True
        lbl_DispToggleCount.Visible = True
        
        For i = 0 To MAX_LANE_COUNT - 1
            cmb_DispShiftSpeed(i).Enabled = True
            cmb_DispShiftSpeed(i).Visible = True
            cmd_NmlShift(i).Enabled = True
            cmd_NmlShift(i).Visible = True
        Next i
    Else
        Cmb_Display_Direct.Visible = False
        'cmb_DisplayMode(0).Visible = False
        cmb_DispToggleTime.Visible = False
        cmb_DispToggleCount.Visible = False
        lbl_DispToggleTime.Visible = False
        lbl_DispToggleCount.Visible = False
        
        For i = 0 To MAX_LANE_COUNT - 1
            cmb_DispShiftSpeed(i).Enabled = False
            cmb_DispShiftSpeed(i).Visible = False
            cmd_NmlShift(i).Enabled = False
            cmd_NmlShift(i).Visible = False
        Next i
    End If
    
    Call Cmb_Display_Direct_Click
    
    
End Sub



Private Sub Cmb_Display_Direct_Click()
    If (Cmb_Display.text = "������(Ǯ�÷�)_FW7") Then
    
        If (Cmb_Display_Direct.text = "����") Then
            lbl_DispToggleCount.Visible = False
            cmb_DispToggleCount.Visible = False
        ElseIf (Cmb_Display_Direct.text = "����") Then
            lbl_DispToggleCount.Visible = True
            cmb_DispToggleCount.Visible = True
        End If
        
    Else
        lbl_DispToggleCount.Visible = False
        cmb_DispToggleCount.Visible = False
    End If
End Sub



Private Sub cmb_PrintModel_Click(Index As Integer)
    If (cmb_PrintModel(Index).text = "NONE") Then
        cmb_PrintPort(Index).Enabled = False
    Else
        cmb_PrintPort(Index).Enabled = True
    End If
End Sub

Private Sub cmd_Button_Click(Index As Integer)
    ' ���� ���� �� ����
    If (Index = 1) Then
        Me.Hide
        Exit Sub
    End If
'On Error Resume Next
    Dim i As Integer
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'APS ����(�ⱸ���������)
    If chk_ApsYN.value = "1" Then
        Glo_ApsYN = "Y"
    Else
        Glo_ApsYN = "N"
    End If
    Glo_Aps_IP = TxtAspIp
    Call Put_Ini("System Config", "APS_YN", Glo_ApsYN)
    Call Put_Ini("System Config", "APS_IP", Glo_Aps_IP)
    
    
    'APS ����(�������������)
    If chk_PreApsYN.value = "1" Then
        Glo_PreApsYN = "Y"
        Glo_Grace_Time = Val(TxtGraceTime)
        Glo_Return_Time = Val(TxtReturnTime)
    Else
        Glo_PreApsYN = "N"
    End If
    Call Put_Ini("System Config", "PreAPS_YN", Glo_PreApsYN)
    Call Put_Ini("System Config", "GRACE_TIME", CStr(Glo_Grace_Time))
    Call Put_Ini("System Config", "RETURN_TIME", CStr(Glo_Return_Time))

'    If (Glo_ApsYN = "Y") Then
'        FrmAccnt.ApsS_sock.Close
'        FrmAccnt.ApsS_sock.Protocol = sckUDPProtocol
'        FrmAccnt.ApsS_sock.LocalPort = 5889
'        FrmAccnt.ApsS_sock.Bind
'    End If
    
    'APS_Port = 5889
    'CMD_Port = 5888
    
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' LPR Board ����
    Call Put_Ini("System Config", "LPRBoard", Glo_LPRBoard)
    
    
    
    ' ������ ���� Set
    '
    If chk_RemoteYN(0).value = "1" Then
        Glo_RemoteR_YN = "Y"
    Else
        Glo_RemoteR_YN = "N"
    End If

    Glo_RemoteR_Port = Val(TxtSvrPort(0))
    Call Put_Ini("System Config", "RemoteR_YN", Glo_RemoteR_YN)
    Call Put_Ini("System Config", "RemoteR_Port", CStr(Glo_RemoteR_Port))

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ������ �۽� Set
    If chk_RemoteYN(1).value = "1" Then
        Glo_RemoteS_YN = "Y"
    Else
        Glo_RemoteS_YN = "N"
    End If

    Glo_RemoteS_IP = Trim(TxtSvrIp(1))
    Glo_RemoteS_Port = Val(TxtSvrPort(1))
    Call Put_Ini("System Config", "RemoteS_YN", CStr(Glo_RemoteS_YN))
    Call Put_Ini("System Config", "RemoteS_IP", CStr(Glo_RemoteS_IP))
    Call Put_Ini("System Config", "RemoteR_Port", CStr(Glo_RemoteS_Port))

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ��ũ�� �� Set
    Select Case Cmb_Window.ListIndex
       Case 0
            Glo_Screen_No = 1
            
'            chk_UseYN(0).value = 1
            chk_UseYN(1).value = 0
            chk_UseYN(2).value = 0
            chk_UseYN(3).value = 0
            chk_UseYN(4).value = 0
            chk_UseYN(5).value = 0
       Case 1
            Glo_Screen_No = 2
            
'            chk_UseYN(0).value = 1
'            chk_UseYN(1).value = 1
            chk_UseYN(2).value = 0
            chk_UseYN(3).value = 0
            chk_UseYN(4).value = 0
            chk_UseYN(5).value = 0
            
            If (CmbScreen(0).ListIndex = 0) Then
                Glo_Screen1 = 1
                Glo_Screen2 = 2
            Else
                Glo_Screen1 = CmbScreen(0).ListIndex
                Glo_Screen2 = CmbScreen(1).ListIndex
            End If
       Case 2
            Glo_Screen_No = 4
            
'            chk_UseYN(0).value = 1
'            chk_UseYN(1).value = 1
'            chk_UseYN(2).value = 1
'            chk_UseYN(3).value = 1
            chk_UseYN(4).value = 0
            chk_UseYN(5).value = 0
            If (CmbScreen(0).ListIndex = 0) Then
                Glo_Screen1 = 1
                Glo_Screen2 = 2
                Glo_Screen3 = 3
                Glo_Screen4 = 4
            Else
                Glo_Screen1 = CmbScreen(0).ListIndex
                Glo_Screen2 = CmbScreen(1).ListIndex
                Glo_Screen3 = CmbScreen(2).ListIndex
                Glo_Screen4 = CmbScreen(3).ListIndex
            End If
        Case 3
            Glo_Screen_No = 6
            
'            chk_UseYN(0).value = 1
'            chk_UseYN(1).value = 1
'            chk_UseYN(2).value = 1
'            chk_UseYN(3).value = 1
'            chk_UseYN(4).value = 1
'            chk_UseYN(5).value = 1
            If (CmbScreen(0).ListIndex = 0) Then
                Glo_Screen1 = 1
                Glo_Screen2 = 2
                Glo_Screen3 = 3
                Glo_Screen4 = 4
                Glo_Screen5 = 5
                Glo_Screen6 = 6
            Else
                Glo_Screen1 = CmbScreen(0).ListIndex
                Glo_Screen2 = CmbScreen(1).ListIndex
                Glo_Screen3 = CmbScreen(2).ListIndex
                Glo_Screen4 = CmbScreen(3).ListIndex
                Glo_Screen5 = CmbScreen(4).ListIndex
                Glo_Screen6 = CmbScreen(5).ListIndex
            End If
    End Select
    
    
    

    '���θ�Ī ����
    Call MainForm_Set_GateName


    If (Frame2(0).Enabled = True) Then
        LANE1_YN = "Y"
    Else
        LANE1_YN = "N"
    End If
    If (Frame2(1).Enabled = True) Then
        LANE2_YN = "Y"
    Else
        LANE2_YN = "N"
    End If
    If (Frame2(2).Enabled = True) Then
        LANE3_YN = "Y"
    Else
        LANE3_YN = "N"
    End If
    If (Frame2(3).Enabled = True) Then
        LANE4_YN = "Y"
    Else
        LANE4_YN = "N"
    End If
    If (Frame2(4).Enabled = True) Then
        LANE5_YN = "Y"
    Else
        LANE5_YN = "N"
    End If
    If (Frame2(5).Enabled = True) Then
        LANE6_YN = "Y"
    Else
        LANE6_YN = "N"
    End If

'''    For i = 0 To Glo_Screen_No - 1
'''        If (Frame2(i).Enabled = True) Then
'''            chk_UseYN(i).value = "1"
'''            cmb_Inout(i).Enabled = True
'''            txt_GateName(i).Enabled = True
'''            CmbScreen(i).Enabled = True
'''            Call MainForm_ChkFreePass(i, True)
'''
'''        Else
'''            chk_UseYN(i).value = "0"
'''            cmb_Inout(i).Enabled = False
'''            txt_GateName(i).Enabled = False
'''            CmbScreen(i).Enabled = False
'''            Call MainForm_ChkFreePass(i, False)
'''        End If
'''    Next
'''    For i = Glo_Screen_No To 5
'''        Frame2(i).Enabled = False
'''        chk_UseYN(i).value = "0"
'''        cmb_Inout(i).Enabled = False
'''        txt_GateName(i).Enabled = False
'''        CmbScreen(i).Enabled = False
'''
'''        Call MainForm_ChkFreePass(i, False)
'''    Next
    


    ' �Ϲ����� ���ⱸ ���о��ְ�, ���κ�ó���� ��ȯ�� - ����
    Call MainForm_SetNormalPass
    ' �Ϲ����� ���ⱸ ���о��ְ�, ���κ�ó���� ��ȯ�� - ��

    ' ���������� ���ⱸ ���о��ְ�, ���κ�ó���� ��ȯ�� - ����
    Call MainForm_SetTaxiPass
    ' ���������� ���ⱸ ���о��ְ�, ���κ�ó���� ��ȯ�� - ��
    
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '�Ĺ�ī�޶� ���� ����
    If (chk_UseYN(0).value = 1 And chk_BackCam_YN(0).value = 1) Then
        Glo_Lane1_Back_YN = "Y"
    Else
        Glo_Lane1_Back_YN = "N"
    End If
    If (chk_UseYN(1).value = 1 And chk_BackCam_YN(1).value = 1) Then
        Glo_Lane2_Back_YN = "Y"
    Else
        Glo_Lane2_Back_YN = "N"
    End If
    If (chk_UseYN(2).value = 1 And chk_BackCam_YN(2).value = 1) Then
        Glo_Lane3_Back_YN = "Y"
    Else
        Glo_Lane3_Back_YN = "N"
    End If
    If (chk_UseYN(3).value = 1 And chk_BackCam_YN(3).value = 1) Then
        Glo_Lane4_Back_YN = "Y"
    Else
        Glo_Lane4_Back_YN = "N"
    End If
    If (chk_UseYN(4).value = 1 And chk_BackCam_YN(4).value = 1) Then
        Glo_Lane5_Back_YN = "Y"
    Else
        Glo_Lane5_Back_YN = "N"
    End If
    If (chk_UseYN(5).value = 1 And chk_BackCam_YN(5).value = 1) Then
        Glo_Lane6_Back_YN = "Y"
    Else
        Glo_Lane6_Back_YN = "N"
    End If
    
    Call Put_Ini("System Config", "LANE1_BACK_YN", Glo_Lane1_Back_YN)
    Call Put_Ini("System Config", "LANE2_BACK_YN", Glo_Lane2_Back_YN)
    Call Put_Ini("System Config", "LANE3_BACK_YN", Glo_Lane3_Back_YN)
    Call Put_Ini("System Config", "LANE4_BACK_YN", Glo_Lane4_Back_YN)
    Call Put_Ini("System Config", "LANE5_BACK_YN", Glo_Lane5_Back_YN)
    Call Put_Ini("System Config", "LANE6_BACK_YN", Glo_Lane6_Back_YN)
    '�Ĺ�ī�޶� ���� ��
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '�湮���� ���� ����
    For i = 0 To 5
        If (Not Glo_FrmGuest(i) Is Nothing) Then '������� �ִٸ�
            'Call Glo_FrmGuest(0).Form_Exit
            Unload Glo_FrmGuest(i)
            Set Glo_FrmGuest(i) = Nothing
        End If
    Next i
    
    If (chk_UseYN(0).value = 1 And LANE1_YN = "Y" And chk_GuestYN(0).value = 1) Then
        Glo_GUEST_LANE1_YN = "Y"
        If (Glo_FrmGuest(0) Is Nothing) Then '������ ���� �ʴٸ�
            Set Glo_FrmGuest(0) = New FormGuest1
            Glo_FrmGuest(0).Show 0
            Call Glo_FrmGuest(0).SetGateNo(0, cmb_PrintModel(0).text, cmb_PrintPort(0).text)
            
        End If
    Else
        Glo_GUEST_LANE1_YN = "N"
    End If
    
    If (chk_UseYN(1).value = 1 And LANE2_YN = "Y" And chk_GuestYN(1).value = 1) Then
        Glo_GUEST_LANE2_YN = "Y"
        If (Glo_FrmGuest(1) Is Nothing) Then '������ ���� �ʴٸ�
            Set Glo_FrmGuest(1) = New FormGuest1
            Glo_FrmGuest(1).Show 0
            Call Glo_FrmGuest(1).SetGateNo(1, cmb_PrintModel(1).text, cmb_PrintPort(1).text)
        End If
    Else
        Glo_GUEST_LANE2_YN = "N"
    End If
    
    If (chk_UseYN(2).value = 1 And LANE3_YN = "Y" And chk_GuestYN(2).value = 1) Then
        Glo_GUEST_LANE3_YN = "Y"
        If (Glo_FrmGuest(2) Is Nothing) Then '������ ���� �ʴٸ�
            Set Glo_FrmGuest(2) = New FormGuest1
            Glo_FrmGuest(2).Show 0
            Call Glo_FrmGuest(2).SetGateNo(2, cmb_PrintModel(2).text, cmb_PrintPort(2).text)
        End If
    Else
        Glo_GUEST_LANE3_YN = "N"
    End If
    
    If (chk_UseYN(3).value = 1 And LANE4_YN = "Y" And chk_GuestYN(3).value = 1) Then
        Glo_GUEST_LANE4_YN = "Y"
        If (Glo_FrmGuest(3) Is Nothing) Then '������ ���� �ʴٸ�
            Set Glo_FrmGuest(3) = New FormGuest1
            Glo_FrmGuest(3).Show 0
            Call Glo_FrmGuest(3).SetGateNo(3, cmb_PrintModel(3).text, cmb_PrintPort(3).text)
        End If
    Else
        Glo_GUEST_LANE4_YN = "N"
    End If
    
    If (chk_UseYN(4).value = 1 And LANE5_YN = "Y" And chk_GuestYN(4).value = 1) Then
        Glo_GUEST_LANE5_YN = "Y"
        If (Glo_FrmGuest(4) Is Nothing) Then '������ ���� �ʴٸ�
            Set Glo_FrmGuest(4) = New FormGuest1
            Glo_FrmGuest(4).Show 0
            Call Glo_FrmGuest(4).SetGateNo(4, cmb_PrintModel(4).text, cmb_PrintPort(4).text)
        End If
    Else
        Glo_GUEST_LANE5_YN = "N"
    End If
    
    If (chk_UseYN(5).value = 1 And LANE6_YN = "Y" And chk_GuestYN(5).value = 1) Then
        Glo_GUEST_LANE6_YN = "Y"
        If (Glo_FrmGuest(5) Is Nothing) Then '������ ���� �ʴٸ�
            Set Glo_FrmGuest(5) = New FormGuest1
            Glo_FrmGuest(5).Show 0
            Call Glo_FrmGuest(5).SetGateNo(5, cmb_PrintModel(5).text, cmb_PrintPort(5).text)
        End If
    Else
        Glo_GUEST_LANE6_YN = "N"
    End If
    
    If (Glo_GUEST_LANE1_YN = "Y" Or Glo_GUEST_LANE2_YN = "Y" Or Glo_GUEST_LANE3_YN = "Y" Or Glo_GUEST_LANE4_YN = "Y" Or Glo_GUEST_LANE5_YN = "Y" Or Glo_GUEST_LANE6_YN = "Y") Then
        Glo_Guest_YN = "Y"
    Else
        Glo_Guest_YN = "N"
    End If
    
    Glo_Guest_Print_Model(0) = cmb_PrintModel(0).text
    Glo_Guest_Print_Model(1) = cmb_PrintModel(1).text
    Glo_Guest_Print_Model(2) = cmb_PrintModel(2).text
    Glo_Guest_Print_Model(3) = cmb_PrintModel(3).text
    Glo_Guest_Print_Model(4) = cmb_PrintModel(4).text
    Glo_Guest_Print_Model(5) = cmb_PrintModel(5).text
    
    Glo_Guest_Print_Port(0) = cmb_PrintPort(0).text
    Glo_Guest_Print_Port(1) = cmb_PrintPort(1).text
    Glo_Guest_Print_Port(2) = cmb_PrintPort(2).text
    Glo_Guest_Print_Port(3) = cmb_PrintPort(3).text
    Glo_Guest_Print_Port(4) = cmb_PrintPort(4).text
    Glo_Guest_Print_Port(5) = cmb_PrintPort(5).text
    
    Call Print_Port_Init(0, Glo_GUEST_LANE1_YN, Glo_Guest_Print_Model(0), Glo_Guest_Print_Port(0))
    Call Print_Port_Init(1, Glo_GUEST_LANE2_YN, Glo_Guest_Print_Model(1), Glo_Guest_Print_Port(1))
    Call Print_Port_Init(2, Glo_GUEST_LANE3_YN, Glo_Guest_Print_Model(2), Glo_Guest_Print_Port(2))
    Call Print_Port_Init(3, Glo_GUEST_LANE4_YN, Glo_Guest_Print_Model(3), Glo_Guest_Print_Port(3))
    Call Print_Port_Init(4, Glo_GUEST_LANE5_YN, Glo_Guest_Print_Model(4), Glo_Guest_Print_Port(4))
    Call Print_Port_Init(5, Glo_GUEST_LANE6_YN, Glo_Guest_Print_Model(5), Glo_Guest_Print_Port(5))
    
    
    Call Put_Ini("System Config", "GUEST1_PRINT_MODEL", Glo_Guest_Print_Model(0))
    Call Put_Ini("System Config", "GUEST2_PRINT_MODEL", Glo_Guest_Print_Model(1))
    Call Put_Ini("System Config", "GUEST3_PRINT_MODEL", Glo_Guest_Print_Model(2))
    Call Put_Ini("System Config", "GUEST4_PRINT_MODEL", Glo_Guest_Print_Model(3))
    Call Put_Ini("System Config", "GUEST5_PRINT_MODEL", Glo_Guest_Print_Model(4))
    Call Put_Ini("System Config", "GUEST6_PRINT_MODEL", Glo_Guest_Print_Model(5))
    
    Call Put_Ini("System Config", "GUEST1_PRINT_PORT", Glo_Guest_Print_Port(0))
    Call Put_Ini("System Config", "GUEST2_PRINT_PORT", Glo_Guest_Print_Port(1))
    Call Put_Ini("System Config", "GUEST3_PRINT_PORT", Glo_Guest_Print_Port(2))
    Call Put_Ini("System Config", "GUEST4_PRINT_PORT", Glo_Guest_Print_Port(3))
    Call Put_Ini("System Config", "GUEST5_PRINT_PORT", Glo_Guest_Print_Port(4))
    Call Put_Ini("System Config", "GUEST6_PRINT_PORT", Glo_Guest_Print_Port(5))
    
    Call Put_Ini("System Config", "GUEST1_YN", Glo_GUEST_LANE1_YN)
    Call Put_Ini("System Config", "GUEST2_YN", Glo_GUEST_LANE2_YN)
    Call Put_Ini("System Config", "GUEST3_YN", Glo_GUEST_LANE3_YN)
    Call Put_Ini("System Config", "GUEST4_YN", Glo_GUEST_LANE4_YN)
    Call Put_Ini("System Config", "GUEST5_YN", Glo_GUEST_LANE5_YN)
    Call Put_Ini("System Config", "GUEST6_YN", Glo_GUEST_LANE6_YN)
    '�湮���� ���� ��
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    
    Call Put_Ini("System Config", "LANE1_YN", LANE1_YN)
    Call Put_Ini("System Config", "LANE2_YN", LANE2_YN)
    Call Put_Ini("System Config", "LANE3_YN", LANE3_YN)
    Call Put_Ini("System Config", "LANE4_YN", LANE4_YN)
    Call Put_Ini("System Config", "LANE5_YN", LANE5_YN)
    Call Put_Ini("System Config", "LANE6_YN", LANE6_YN)
    
    'Call Put_Ini("System Config", "TAXI_IN_YN", Glo_TAXI_IN_YN)
    'Call Put_Ini("System Config", "TAXI_OUT_YN", Glo_TAXI_OUT_YN)

    Call Put_Ini("System Config", "LANE1_ȭ����ġ", CStr(Glo_Screen1))
    Call Put_Ini("System Config", "LANE2_ȭ����ġ", CStr(Glo_Screen2))
    Call Put_Ini("System Config", "LANE3_ȭ����ġ", CStr(Glo_Screen3))
    Call Put_Ini("System Config", "LANE4_ȭ����ġ", CStr(Glo_Screen4))
    Call Put_Ini("System Config", "LANE5_ȭ����ġ", CStr(Glo_Screen5))
    Call Put_Ini("System Config", "LANE6_ȭ����ġ", CStr(Glo_Screen6))
    
    Call Put_Ini("System Config", "Screen_No", CStr(Glo_Screen_No))
    
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'If (cmb_Board.text = "�����" Or cmb_Board.text = "�ڵ��̳�") Then
        ' �����ġ ���
        Select Case Cmb_Display.ListIndex
           Case 0
                Glo_Display = "������"
           Case 1
                Glo_Display = "FND"
           Case 2
                Glo_Display = "������(Ǯ�÷�)"
           Case 3
                Glo_Display = "������(Ǯ�÷�)_FW7"
        End Select
        
        Glo_Display_Direct = Cmb_Display_Direct.text '����
        Glo_Emerg_Vertical_ToggleTime = cmb_DispToggleTime.text '��������� �����ð�(s)
        Glo_Emerg_Vertical_ToggleCount = cmb_DispToggleCount.text '���Ƚ��

        Call Put_Ini("System Config", "Display", CStr(Glo_Display))
        Call Put_Ini("System Config", "Display_Direct", Glo_Display_Direct)
        adoConn.Execute "UPDATE tb_config set Content = '" & Glo_Emerg_Vertical_ToggleTime & "' WHERE NAME = 'Disp_Vertical_ToggleTime'"
        adoConn.Execute "UPDATE tb_config set Content = '" & Glo_Emerg_Vertical_ToggleCount & "' WHERE NAME = 'Disp_Vertical_ToggleCount'"
    'End If
    

        
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'LPR Set
    For i = 0 To MAX_LANE_COUNT - 1
        Call cmd_OK_Click(i)
    Next i
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'    'Sever ReStart(TCP)
'    If (LANE1_LPRMode = "0") Then
'        Call Server.StopServer
'        Call Server.StartServer(Server_Port, Server.ServerIP)
'    End If

    
    
    Call SetCommunication
    
  
    On Error Resume Next
    Select Case Cmb_Window.ListIndex
           Case 0 '����ȭ��
                FrmG1.Show 0
                Jung.Hide
                FrmG4Mini.Hide
                FrmG6_23.Hide
'                If (Glo_ApsYN = "Y") Then
'                    FrmG1.Lblbutton(7).Visible = True
'                    FrmG1.Imgbutton(7).Visible = True
'                    FrmG1.Lblbutton(8).Visible = True
'                    FrmG1.Imgbutton(8).Visible = True
'                Else
'                    FrmG1.Lblbutton(7).Visible = False
'                    FrmG1.Imgbutton(7).Visible = False
'                    FrmG1.Lblbutton(8).Visible = False
'                    FrmG1.Imgbutton(8).Visible = False
'                End If
                
           Case 1 '2ȭ��
                FrmG1.Hide
                Jung.Show 0
                FrmG4Mini.Hide
                FrmG6_23.Hide
'                If (Glo_ApsYN = "Y") Then
'                    Jung.Lblbutton(7).Visible = True
'                    Jung.Imgbutton(7).Visible = True
'                    Jung.Lblbutton(8).Visible = True
'                    Jung.Imgbutton(8).Visible = True
'                Else
'                    Jung.Lblbutton(7).Visible = False
'                    Jung.Imgbutton(7).Visible = False
'                    Jung.Lblbutton(8).Visible = False
'                    Jung.Imgbutton(8).Visible = False
'                End If
           Case 2 '4ȭ��
                Jung.Hide
                FrmG1.Hide
                FrmG4Mini.Show 0
                FrmG6_23.Hide
'                If (Glo_ApsYN = "Y") Then
'                    FrmG4Mini.Lblbutton(7).Visible = True
'                    FrmG4Mini.Imgbutton(7).Visible = True
'                    FrmG4Mini.Lblbutton(8).Visible = True
'                    FrmG4Mini.Imgbutton(8).Visible = True
'                Else
'                    FrmG4Mini.Lblbutton(7).Visible = False
'                    FrmG4Mini.Imgbutton(7).Visible = False
'                    FrmG4Mini.Lblbutton(8).Visible = False
'                    FrmG4Mini.Imgbutton(8).Visible = False
'                End If
            Case 3 '6ȭ��
                Jung.Hide
                FrmG1.Hide
                FrmG4Mini.Hide
                FrmG6_23.Show 0
'                If (Glo_ApsYN = "Y") Then
'                    FrmG6_23.cmd_menu(8).Visible = True
'                    'FrmG6_23.cmd_menu(8).Enabled = True
'                Else
'                    FrmG6_23.cmd_menu(8).Visible = False
'                    'FrmG6_23.cmd_menu(8).Enabled = False
'                End If
            
    
    End Select
    Call frmLogin.ShowMenu(Glo_Login_ID, Glo_Login_PW)
    Me.Show 0
    'TxtSvrIp(1).Refresh
    
'    Me.Hide

    'FrmTcpServer.ListData.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & "[ȯ�漳�� ����]", 0
    Call DataLogger("[ȯ�漳�� ����]")

End Sub

' �Ϲ����� �ڵ� ���� ����
Private Sub MainForm_SetNormalPass()

    Dim i As Integer
    Dim sLaneUse As String
    Dim sLaneName As String
    Dim sNormalUse As String
    For i = 0 To Glo_Screen_No - 1

        If (i < Glo_Screen_No) Then
            If (i = 0) Then
                sLaneUse = LANE1_YN:            sLaneName = LANE1_Name:            sNormalUse = Glo_FreePassLane1_YN
            ElseIf (i = 1) Then
                sLaneUse = LANE2_YN:            sLaneName = LANE2_Name:            sNormalUse = Glo_FreePassLane2_YN
            ElseIf (i = 2) Then
                sLaneUse = LANE3_YN:            sLaneName = LANE3_Name:            sNormalUse = Glo_FreePassLane3_YN
            ElseIf (i = 3) Then
                sLaneUse = LANE4_YN:            sLaneName = LANE4_Name:            sNormalUse = Glo_FreePassLane4_YN
            ElseIf (i = 4) Then
                sLaneUse = LANE5_YN:            sLaneName = LANE5_Name:            sNormalUse = Glo_FreePassLane5_YN
            ElseIf (i = 5) Then
                sLaneUse = LANE6_YN:            sLaneName = LANE6_Name:            sNormalUse = Glo_FreePassLane6_YN
            End If
            
            If (Glo_Screen_No = 1) Then
                Call Chk_NormalPassEnable(FrmG1, sLaneUse, sNormalUse, i, sLaneName)
            ElseIf (Glo_Screen_No = 2) Then
                Call Chk_NormalPassEnable(Jung, sLaneUse, sNormalUse, i, sLaneName)
            ElseIf (Glo_Screen_No = 4) Then
                Call Chk_NormalPassEnable(FrmG4Mini, sLaneUse, sNormalUse, i, sLaneName)
            ElseIf (Glo_Screen_No = 6) Then
                Call Chk_NormalPassEnable(FrmG6_23, sLaneUse, sNormalUse, i, sLaneName)
            End If

        End If
    Next

End Sub

' �������� �ڵ� ���� ����
Private Sub MainForm_SetTaxiPass()

    Dim i As Integer
    Dim sLaneUse As String
    Dim sLaneName As String
    Dim sTaxiUse As String
    For i = 0 To Glo_Screen_No - 1

        If (i < Glo_Screen_No) Then
            If (i = 0) Then
                sLaneUse = LANE1_YN:            sLaneName = LANE1_Name:            sTaxiUse = Glo_TAXI1_YN
            ElseIf (i = 1) Then
                sLaneUse = LANE2_YN:            sLaneName = LANE2_Name:            sTaxiUse = Glo_TAXI2_YN
            ElseIf (i = 2) Then
                sLaneUse = LANE3_YN:            sLaneName = LANE3_Name:            sTaxiUse = Glo_TAXI3_YN
            ElseIf (i = 3) Then
                sLaneUse = LANE4_YN:            sLaneName = LANE4_Name:            sTaxiUse = Glo_TAXI4_YN
            ElseIf (i = 4) Then
                sLaneUse = LANE5_YN:            sLaneName = LANE5_Name:            sTaxiUse = Glo_TAXI5_YN
            ElseIf (i = 5) Then
                sLaneUse = LANE6_YN:            sLaneName = LANE6_Name:            sTaxiUse = Glo_TAXI6_YN
            End If
            
            If (Glo_Screen_No = 1) Then
                Call Chk_TaxiPassEnable(FrmG1, sLaneUse, sTaxiUse, i, sLaneName)
            ElseIf (Glo_Screen_No = 2) Then
                Call Chk_TaxiPassEnable(Jung, sLaneUse, sTaxiUse, i, sLaneName)
            ElseIf (Glo_Screen_No = 4) Then
                Call Chk_TaxiPassEnable(FrmG4Mini, sLaneUse, sTaxiUse, i, sLaneName)
            ElseIf (Glo_Screen_No = 6) Then
                Call Chk_TaxiPassEnable(FrmG6_23, sLaneUse, sTaxiUse, i, sLaneName)
            End If

        End If
    Next
End Sub

Private Sub MainForm_ChkFreePass(ByVal iIdx As Integer, ByVal bVal As Boolean)
    If (iIdx < Glo_Screen_No) Then
        If (Glo_Screen_No = 1) Then
            Call Chk_FreePassEnable(FrmG1, iIdx, bVal)
        ElseIf (Glo_Screen_No = 2) Then
            Call Chk_FreePassEnable(Jung, iIdx, bVal)
        ElseIf (Glo_Screen_No = 4) Then
            Call Chk_FreePassEnable(FrmG4Mini, iIdx, bVal)
        ElseIf (Glo_Screen_No = 6) Then
            Call Chk_FreePassEnable(FrmG6_23, iIdx, bVal)
        End If
    End If

End Sub

Private Sub MainForm_Set_GateName()
    Dim i As Integer

    If (Glo_Screen_No = 1) Then
            For i = 0 To Glo_Screen_No - 1
                Call FrmG1.Set_GateName(i, txt_GateName(i).text)
            Next
    ElseIf (Glo_Screen_No = 2) Then
            For i = 0 To Glo_Screen_No - 1
                Call Jung.Set_GateName(i, txt_GateName(i).text)
            Next
    ElseIf (Glo_Screen_No = 4) Then
            For i = 0 To Glo_Screen_No - 1
                Call FrmG4Mini.Set_GateName(i, txt_GateName(i).text)
            Next
    ElseIf (Glo_Screen_No = 6) Then
            For i = 0 To Glo_Screen_No - 1
                'Debug.Print txt_GateName(i).text
                Call FrmG6_23.Set_GateName(i, txt_GateName(i).text)
            Next
    End If
End Sub


' �Ա��� ���� ������� üũ
Public Function Get_bUseInGate() As Boolean
    
    Dim i As Integer
    Dim bUseGate As Boolean
    bUseGate = False
    
    For i = 0 To MAX_LANE_COUNT - 1
        If (Frame2(i).Enabled = True) Then  ' ���� Ȱ��ȭ
            If (cmb_Inout(i).ListIndex = 0) Then '�Ա�
                bUseGate = True
                Exit For
            End If
        End If
    Next
    
    Get_bUseInGate = bUseGate
End Function

' �ⱸ�� ���� ������� üũ
Public Function Get_bUseOutGate() As Boolean
    
    Dim i As Integer
    Dim bUseGate As Boolean
    bUseGate = False
    
    For i = 0 To MAX_LANE_COUNT - 1
        If (Frame2(i).Enabled = True) Then  ' ���� Ȱ��ȭ
            
            If (cmb_Inout(i).ListIndex = 1) Then '�ⱸ
                bUseGate = True
                Exit For
            End If
        End If
    Next
    
    Get_bUseOutGate = bUseGate
End Function

Private Sub cmd_DeviceReset_Click(Index As Integer)
    If (Glo_LPRBoard = "�ڵ��̳�") Then
        Reset_sock(Index).SendData ("RESET") '���� S/W ����
        Call None_Delay_Time(0.1)
        Call DataLogger("[DEVICE RESET]  Target Gate = " & Index)
        adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('System', 'HOST','LANE" & Index + 1 & " �ڵ��̳� ����!!',''," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
    Else
        Call DataLogger("[DEVICE RESET]  ����̽� ���� �� ���߽��ϴ�")
    End If
End Sub

Private Sub cmd_DispReset_Click(Index As Integer)
    Call GL_Display_PowerOFF(Index)
    Call Delay_Time(3)
    Call GL_Display_PowerON(Index)
End Sub


Private Sub cmd_GateTestDown_Click(Index As Integer)

    GlO_TcpDataGate = Chr$(2) & "GATE DOWN" & Chr$(3) '���ܱ����(���ܱ� ��Ʈ�ѷ� ��������:FPtech�� �������� ������)(Test:Debug��)
    Me.Gate1_sock(Index).SendData (GlO_TcpDataGate)
    If (Index = 0) Then
        Call DataLogger("[GATE DOWN UDP ����]  IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort)
    ElseIf (Index = 1) Then
        Call DataLogger("[GATE DOWN UDP ����]  IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort)
    ElseIf (Index = 2) Then
        Call DataLogger("[GATE DOWN UDP ����]  IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_RelayPort)
    ElseIf (Index = 3) Then
        Call DataLogger("[GATE DOWN UDP ����]  IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort)
    ElseIf (Index = 4) Then
        Call DataLogger("[GATE DOWN UDP ����]  IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_RelayPort)
    ElseIf (Index = 5) Then
        Call DataLogger("[GATE DOWN UDP ����]  IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_RelayPort)
    End If
End Sub

Private Sub cmd_NmlShift_Click(Index As Integer)
    Dim upColor As Byte
    Dim downColor As Byte
    
    Select Case cmb_Disp1(Index).text
        Case "��"
            upColor = enumDIS_COLORs.eRED
        Case "Ȳ"
            upColor = enumDIS_COLORs.eYellow
        Case "��"
            upColor = enumDIS_COLORs.eGreen
        Case "��"
            upColor = enumDIS_COLORs.eBLUE
        Case "��"
            upColor = enumDIS_COLORs.eWINE
        Case "��"
            upColor = enumDIS_COLORs.eSKY
        Case "��"
            upColor = enumDIS_COLORs.eWHITE
    End Select
    Select Case cmb_Disp2(Index).text
        Case "��"
            downColor = enumDIS_COLORs.eRED
        Case "Ȳ"
            downColor = enumDIS_COLORs.eYellow
        Case "��"
            downColor = enumDIS_COLORs.eGreen
        Case "��"
            downColor = enumDIS_COLORs.eBLUE
        Case "��"
            downColor = enumDIS_COLORs.eWINE
        Case "��"
            downColor = enumDIS_COLORs.eSKY
        Case "��"
            downColor = enumDIS_COLORs.eWHITE
    End Select
    
    '�Ϲݹ��� �̵� �Ǵ� ����
    If (cmd_NmlShift(Index).Caption = "�̵�") Then
        cmd_NmlShift(Index).Caption = "����"
        cmb_DispShiftSpeed(Index).Visible = False
        Glo_LANE_DISP_NML_SHIFT(Index) = enumDISP_NML_SHIFT.eSTOP
            
        If (Glo_Display_Direct = "����") Then
            txt_Disp1(Index) = LeftH(txt_Disp1(Index), Glo_DISP_COL * 2) '���� �������� 6���̹Ƿ� 12����(6x2) ������
            txt_Disp2(Index) = LeftH(txt_Disp2(Index), Glo_DISP_COL * 2) '���� �������� 6���̹Ƿ� 12����(6x2) ������
        
        ElseIf (Glo_Display_Direct = "����") Then
            txt_Disp1(Index) = Left(txt_Disp1(Index), Glo_DISP_COL) '���� �������� 6���̹Ƿ� 6���� ������
            txt_Disp2(Index) = Left(txt_Disp2(Index), Glo_DISP_COL) '���� �������� 6���̹Ƿ� 6���� ������
        End If
        
    ElseIf (cmd_NmlShift(Index).Caption = "����") Then
        cmd_NmlShift(Index).Caption = "�̵�"
        cmb_DispShiftSpeed(Index).Visible = True
        Glo_LANE_DISP_NML_SHIFT(Index) = enumDISP_NML_SHIFT.eSHIFT
    End If
    Select Case Index
        Case 0
            Call Put_Ini("System Config", "LANE1_DispShift", CStr(Glo_LANE_DISP_NML_SHIFT(Index)))
        Case 1
            Call Put_Ini("System Config", "LANE2_DispShift", CStr(Glo_LANE_DISP_NML_SHIFT(Index)))
        Case 2
            Call Put_Ini("System Config", "LANE3_DispShift", CStr(Glo_LANE_DISP_NML_SHIFT(Index)))
        Case 3
            Call Put_Ini("System Config", "LANE4_DispShift", CStr(Glo_LANE_DISP_NML_SHIFT(Index)))
        Case 4
            Call Put_Ini("System Config", "LANE5_DispShift", CStr(Glo_LANE_DISP_NML_SHIFT(Index)))
        Case 5
            Call Put_Ini("System Config", "LANE6_DispShift", CStr(Glo_LANE_DISP_NML_SHIFT(Index)))
    End Select
    
    'Display Nomal Save
    'If (Glo_LPRBoard = "�����") Then
    If (Glo_Display = "������" Or Glo_Display = "������(Ǯ�÷�)") Then
        Call DataLogger("[DISPLAY Nomal Shift]  Target Gate = " & Index)
        Call GL_Nomal(txt_Disp1(Index), txt_Disp2(Index), 129, 70, 0, cmb_Disp1(Index).ListIndex, cmb_Disp2(Index).ListIndex, Index) 'OLD ������ �߿���, �������
        
        
    'ElseIf (Glo_LPRBoard = "�ڵ��̳�") Then
    ElseIf (Glo_Display = "������(Ǯ�÷�)_FW7") Then
        Call DataLogger("[DISPLAY Nomal Shift]  Target Gate = " & Index)
        'Call GL_Nomal(txt_Disp1(Index), txt_Disp2(Index), 129, 70, 0, cmb_Disp1(Index).ListIndex, cmb_Disp2(Index).ListIndex, Index) 'OLD ������ �߿���, �������

        If (Glo_Display_Direct = "����") Then
            Call GL_Nomal_Horizontal(txt_Disp1(Index), txt_Disp2(Index), 129, cmb_DispShiftSpeed(Index).text * 10, 0, upColor, downColor, Index, Glo_LANE_DISP_NML_SHIFT(Index)) '������ ����ǥ��(DABIT ������ ��� ��������:HEX), �������
        Else
            Call GL_Nomal_Vertical(txt_Disp1(Index), txt_Disp2(Index), 129, cmb_DispShiftSpeed(Index).text * 10, 0, upColor, downColor, Index, Glo_LANE_DISP_NML_SHIFT(Index)) '������ ����ǥ��(NEW ������ �߿���), �������
        End If
    Else
        Call DataLogger("DISPLAY Nomal Shift TEST Error: " & Glo_LPRBoard)
        Exit Sub
    End If
    
    Call SaveNmlMsg(Index)
    
End Sub

Private Sub Command1_Click()
    FormSound.Show 1
End Sub



Private Sub cmd_Certify_Click()

    Dim rs As ADODB.Recordset
    Dim qry As String
    Dim bQryResult As Boolean
    Dim sIP, sMac As String
    
    Call GetClientIP(Glo_IPAddr)
    Call GetClientMac(Glo_MacAddr)
    Call GetClienKey(Glo_PhyHDDKey)

'    If (cmd_Certify.Caption = "����") Then
    If (Glo_Certify = enumCertify.eCertNoTry) Then '����������
        
        
        
        Set rs = New ADODB.Recordset
        qry = "SELECT LockDate, UnLockDate FROM tb_Certify WHERE HASHCODE = '" & Glo_PhyHDDKey & "' "
    
        If (DataBaseQuery(rs, adoConn, qry, NWERR_GATE_STAY) = False) Then
            Call DebugLogger("[CERTIFY]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
            Exit Sub
        End If
        
        If rs.EOF Then
            bQryResult = DataBaseQueryExec(adoConn, "INSERT INTO tb_certify (LockDate, UnLockDate, IP, Mac, Hashcode, Memo ) VALUES ('" & Format(Now, "yyyy-mm-dd") & "','', '" & Glo_IPAddr & "', '" & Glo_MacAddr & "', '" & Glo_PhyHDDKey & "', '')", NWERR_GATE_STAY, 0)
            If (bQryResult = False) Then
                DataLogger ("[CERTIFY]    " & "������ �������� �ʾҽ��ϴ�. �ٽ� �õ����ּ���.")
                Set rs = Nothing
                Exit Sub
            End If
            
        Else
            bQryResult = DataBaseQueryExec(adoConn, "UPDATE tb_certify SET LockDate='" & Format(Now, "yyyy-mm-dd") & "' WHERE HASHCODE = '" & Glo_PhyHDDKey & "' ", NWERR_GATE_STAY)
            If (bQryResult = False) Then
                DataLogger ("[CERTIFY]    " & "������ �������� �ʾҽ��ϴ�. �ٽ� �õ����ּ���.")
                Set rs = Nothing
                Exit Sub
            End If
            
            
            
        End If
    
    
        cmd_Certify.Caption = "�����ʿ�"
        cmd_Certify.ToolTipText = "�ݵ�� ������ �������� ������������. ���������ĺ��� ���ܱⰡ ���������� �ʽ��ϴ�"
        txt_CertifyKey.Visible = True
        lbl_CertifyLimitDate.Caption = "����Ⱓ:" & DateAdd("m", Glo_Cert_Month, Format(Now, "yyyy-mm-dd"))
        lbl_CertifyLimitDate.Visible = True

        Glo_Cert_LimitDate = DateAdd("m", Glo_Cert_Month, Format(Now, "yyyy-mm-dd")) '������
        Glo_Cert_NoticeSDate = DateAdd("m", Glo_Cert_Month - 1, Format(Now, "yyyy-mm-dd")) '����Ⱓ �ȳ� ������(������ 1���� ��)
        Glo_Certify = enumCertify.eCertTry
            

'    ElseIf (cmd_Certify.Caption = "�����ʿ�") Then
    ElseIf (Glo_Certify = enumCertify.eCertTry) Then

        If (Len(txt_CertifyKey.text) = 0) Then
            Call DataLogger("����Ű �Է����ּ���")

        Else
            If (txt_CertifyKey.text = "admin" & Format(Now, "ddmm")) Then 'admin�Ͽ�
                bQryResult = DataBaseQueryExec(adoConn, "UPDATE tb_certify SET UnLockDate='" & Format(Now, "yyyy-mm-dd") & "' WHERE HASHCODE = '" & Glo_PhyHDDKey & "' ", NWERR_GATE_STAY)
                If (bQryResult = False) Then
                    DataLogger ("[Certify_Click]    " & "���������� �������� �ʾҽ��ϴ�. �ٽ� �õ����ּ���.")
                    Set rs = Nothing
                    Exit Sub
                End If

                cmd_Certify.Caption = "�����Ϸ�"
                cmd_Certify.ToolTipText = ""
                cmd_Certify.Visible = True
                cmd_Certify.Enabled = False
                txt_CertifyKey.Visible = False
                lbl_CertifyLimitDate.Visible = False
                txt_CertifyKey.text = ""
                Glo_Certify = enumCertify.eCertOK
                Call DataLogger("��������")

            ElseIf (txt_CertifyKey.text = "jawootek" & Format(Now, "ddmm")) Then 'admin�Ͽ�
                 Glo_COMPANY = "(��)�ڿ���"
            Else
                txt_CertifyKey.text = ""
                Call DataLogger("��������!!")
            End If
        End If
    End If
    
End Sub



'����� �˸� ��ư
Private Sub Command10_Click()
    FormMobile.Show 1
End Sub


Private Sub Command2_Click()

    If (Trim(txt_Vendor) > 0) Then
        If (Trim(txt_SiteName) > 0) Then
        
            Call SendCertPacket("REQ_SITEREG_" & Glo_IPAddr & "_" & Glo_MacAddr & "_" & Glo_PhyHDDKey & "_" & Trim(txt_Vendor) & "_" & Trim(txt_SiteName))
        Else
            txt_SiteName.text = ""
            txt_SiteName.SetFocus
        End If
    Else
        txt_Vendor.text = ""
        txt_Vendor.SetFocus
    End If
    
    
End Sub

Private Sub Command9_Click()
    FormIPCamera.Show 1
End Sub

Private Sub DeviceR_sock_DataArrival(ByVal bytesTotal As Long)
    Dim sdata As String
    Dim sStrLine() As String
    Dim gateNo As Integer
    Dim CmdR As String
    
On Error GoTo Err_p

    If (bytesTotal > 500) Then
        Exit Sub
    End If
    
    DeviceR_sock.GetData sdata, , bytesTotal
    sdata = "" & sdata
    
'    If (sdata = "") Then
'        Exit Sub
'    End If
'    Debug.Print sdata
    
'    sStrLine() = Split(sdata, "_")
'    GateNo = sStrLine(0)
'    CmdR = sStrLine(1)
    
    Call DataLogger("DeviceR_sock UDP Port " & FormatNumber(bytesTotal, 0, , , vbTrue) & " bytes recieved." & "    " & sdata)

Exit Sub

Err_p:
    Call DataLogger(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    [DeviceR_sock Err]  " & Err.Description)
End Sub

Private Sub DeviceR_sock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Call DataLogger(" [DeviceR_sock UDP Error]  " & Description)
End Sub

Private Sub Form_Activate()
frontno = 10


    Dim Port As Integer
    Dim i As Integer
    Dim bScrNoChk As Boolean
    
'On Error GoTo Err_Proc
    Call InitializeCriticalSection(Glo_CS)
    
    
    Left = (Screen.width - width) / 2   ' ���� ���η� �߾ӿ� �����ϴ�.
    Top = (Screen.height - height) / 2   ' ���� ���η� �߾ӿ� �����ϴ�.
    
    Call Certify 'ȣ��Ʈ ����ó��
    
    txtIP = Server.ServerIP
    txtPort = Server_Port



    'APS ����
    Glo_ApsYN = Get_Ini("System Config", "APS_YN", "N")
    Glo_Aps_IP = Get_Ini("System Config", "APS_IP", "127.0.0.1")
    TxtAspIp = Glo_Aps_IP
    If (Glo_ApsYN = "Y") Then
        chk_ApsYN.value = 1
        TxtAspIp.BackColor = &H80000005
        Frame7.Enabled = True
    Else
        chk_ApsYN.value = 0
        TxtAspIp.BackColor = &HE0E0E0
        Frame7.Enabled = False
    End If
    If (Glo_PreApsYN = "Y") Then
        chk_PreApsYN.value = 1
        TxtGraceTime.BackColor = &H80000005
        TxtReturnTime.BackColor = &H80000005
        Frame9.Enabled = True
    Else
        chk_PreApsYN.value = 0
        TxtGraceTime.BackColor = &HE0E0E0
        TxtReturnTime.BackColor = &HE0E0E0
        Frame9.Enabled = False
    End If
    
    
    If (Glo_ApsYN = "Y") Then
        Aps_UDP.Close
        Aps_UDP.Protocol = sckUDPProtocol
        Aps_UDP.LocalPort = 5887
        Aps_UDP.Bind
    End If
    
    
    
    If (Glo_LPRBoard = "�����" Or Glo_LPRBoard = "�ڵ��̳�") Then
        cmb_Board.text = Glo_LPRBoard
    Else
        DataLogger ("LPR Board �������: " & Glo_LPRBoard)
    End If
    
    
    Glo_RemoteS_YN = Get_Ini("System Config", "RemoteS_YN", "N")
    Glo_RemoteS_IP = Get_Ini("System Config", "RemoteS_IP", "127.0.0.1")
    Glo_RemoteS_Port = Val(Get_Ini("System Config", "RemoteS_Port", "7000"))
    
    Glo_RemoteR_YN = Get_Ini("System Config", "RemoteR_YN", "N")
    Glo_RemoteR_Port = Val(Get_Ini("System Config", "RemoteR_Port", "7000"))
    
    
    Glo_FreepassS_YN = Get_Ini("System Config", "FreepassS_YN", "N")
    Glo_FreepassS_IP = Get_Ini("System Config", "FreepassS_IP", "127.0.0.1")
    Glo_FreepassS_Port = Val(Get_Ini("System Config", "FreepassS_Port", "18280"))
    Glo_FreepassR_YN = Get_Ini("System Config", "FreepassR_YN", "N")
    Glo_FreepassR_Port = Val(Get_Ini("System Config", "FreepassR_Port", "18280"))
    
    
    '�Ĺ�ī�޶� ���� �ε�
    If (Glo_Lane1_Back_YN = "Y") Then
        chk_BackCam_YN(0).value = 1
    Else
        chk_BackCam_YN(0).value = 0
    End If
    If (Glo_Lane2_Back_YN = "Y") Then
        chk_BackCam_YN(1).value = 1
    Else
        chk_BackCam_YN(1).value = 0
    End If
    If (Glo_Lane3_Back_YN = "Y") Then
        chk_BackCam_YN(2).value = 1
    Else
        chk_BackCam_YN(2).value = 0
    End If
    If (Glo_Lane4_Back_YN = "Y") Then
        chk_BackCam_YN(3).value = 1
    Else
        chk_BackCam_YN(3).value = 0
    End If
    If (Glo_Lane5_Back_YN = "Y") Then
        chk_BackCam_YN(4).value = 1
    Else
        chk_BackCam_YN(4).value = 0
    End If
    If (Glo_Lane6_Back_YN = "Y") Then
        chk_BackCam_YN(5).value = 1
    Else
        chk_BackCam_YN(5).value = 0
    End If
    '�Ĺ�ī�޶� ���� �ε�
    
    If (Glo_GUEST_LANE1_YN = "Y") Then
        chk_GuestYN(0).value = 1
    Else
        chk_GuestYN(0).value = 0
    End If
    If (Glo_GUEST_LANE2_YN = "Y") Then
        chk_GuestYN(1).value = 1
    Else
        chk_GuestYN(1).value = 0
    End If
    If (Glo_GUEST_LANE3_YN = "Y") Then
        chk_GuestYN(2).value = 1
    Else
        chk_GuestYN(2).value = 0
    End If
    If (Glo_GUEST_LANE4_YN = "Y") Then
        chk_GuestYN(3).value = 1
    Else
        chk_GuestYN(3).value = 0
    End If
    If (Glo_GUEST_LANE5_YN = "Y") Then
        chk_GuestYN(4).value = 1
    Else
        chk_GuestYN(4).value = 0
    End If
    If (Glo_GUEST_LANE6_YN = "Y") Then
        chk_GuestYN(5).value = 1
    Else
        chk_GuestYN(5).value = 0
    End If
    If (Glo_GUEST_LANE1_YN = "Y" Or Glo_GUEST_LANE2_YN = "Y" Or Glo_GUEST_LANE3_YN = "Y" Or Glo_GUEST_LANE4_YN = "Y" Or Glo_GUEST_LANE5_YN = "Y" Or Glo_GUEST_LANE6_YN = "Y") Then
        Glo_Guest_YN = "Y"
    Else
        Glo_Guest_YN = "N"
    End If
    
    
    

    If Glo_RemoteR_YN = "Y" Then
        chk_RemoteYN(0).value = 1
        Frame1(0).Enabled = True
    Else
        chk_RemoteYN(0).value = 0
        Frame1(0).Enabled = False
    End If
    

    If Glo_RemoteS_YN = "Y" Then
        chk_RemoteYN(1).value = 1
        TxtSvrIp(1).BackColor = &H80000005
        Frame1(1).Enabled = True
    Else
        chk_RemoteYN(1).value = 0
        TxtSvrIp(1).BackColor = &HE0E0E0
        Frame1(1).Enabled = False
    End If
    TxtSvrPort(0) = Glo_RemoteR_Port
    TxtSvrIp(1) = Glo_RemoteS_IP
    TxtSvrPort(1) = Glo_RemoteS_Port
    
    
    If (Glo_Display = "������(Ǯ�÷�)" Or Glo_Display = "������(Ǯ�÷�)_FW7") Then

        For i = 0 To MAX_LANE_COUNT - 1
            cmb_Disp1(i).Clear
            cmb_Disp1(i).AddItem "��"
            cmb_Disp1(i).AddItem "��"
            cmb_Disp1(i).AddItem "Ȳ"
            cmb_Disp1(i).AddItem "��"
            cmb_Disp1(i).AddItem "��"
            cmb_Disp1(i).AddItem "��"
            cmb_Disp1(i).AddItem "��"
            
            cmb_Disp2(i).Clear
            cmb_Disp2(i).AddItem "��"
            cmb_Disp2(i).AddItem "��"
            cmb_Disp2(i).AddItem "Ȳ"
            cmb_Disp2(i).AddItem "��"
            cmb_Disp2(i).AddItem "��"
            cmb_Disp2(i).AddItem "��"
            cmb_Disp2(i).AddItem "��"
        Next i
    Else
        For i = 0 To MAX_LANE_COUNT - 1
            cmb_Disp1(i).Clear
            cmb_Disp1(i).AddItem "��"
            cmb_Disp1(i).AddItem "��"
            cmb_Disp1(i).AddItem "Ȳ"

            cmb_Disp2(i).Clear
            cmb_Disp2(i).AddItem "��"
            cmb_Disp2(i).AddItem "��"
            cmb_Disp2(i).AddItem "Ȳ"
        Next i
    End If
    
    
    '��޹��� ���Ÿ��(������ȯ�ð�)
    cmb_DispToggleTime.Clear
    cmb_DispToggleTime.AddItem enumDISP_EMG_TIME.e1sec
    cmb_DispToggleTime.AddItem enumDISP_EMG_TIME.e2sec
    cmb_DispToggleTime.AddItem enumDISP_EMG_TIME.e3sec
    cmb_DispToggleTime.AddItem enumDISP_EMG_TIME.e4sec
    cmb_DispToggleTime.AddItem enumDISP_EMG_TIME.e5sec
    cmb_DispToggleTime.AddItem enumDISP_EMG_TIME.e6sec
    cmb_DispToggleTime.AddItem enumDISP_EMG_TIME.e7sec
    cmb_DispToggleTime.AddItem enumDISP_EMG_TIME.e8sec
    cmb_DispToggleTime.AddItem enumDISP_EMG_TIME.e9sec
    cmb_DispToggleTime.AddItem enumDISP_EMG_TIME.e10sec
    If ((Glo_Emerg_Vertical_ToggleTime >= enumDISP_EMG_TIME.e1sec) And (Glo_Emerg_Vertical_ToggleTime <= enumDISP_EMG_TIME.e10sec)) Then
        cmb_DispToggleTime.text = Glo_Emerg_Vertical_ToggleTime '��������� �����ð�(s)
    Else
        cmb_DispToggleTime.text = enumDISP_EMG_TIME.e3sec '�⺻�� 3000ms
    End If
    
    '��޹��� ǥ��Ƚ��
    cmb_DispToggleCount.Clear
    cmb_DispToggleCount.AddItem "1"
    cmb_DispToggleCount.AddItem "2"
    cmb_DispToggleCount.AddItem "3"
    cmb_DispToggleCount.AddItem "4"
    cmb_DispToggleCount.AddItem "5"
    cmb_DispToggleCount.AddItem "6"
    cmb_DispToggleCount.AddItem "7"
    cmb_DispToggleCount.AddItem "8"
    cmb_DispToggleCount.AddItem "9"
    cmb_DispToggleCount.AddItem "10"
    If ((Glo_Emerg_Vertical_ToggleCount >= "1") And (Glo_Emerg_Vertical_ToggleCount <= "10")) Then
        cmb_DispToggleCount.text = CStr(Glo_Emerg_Vertical_ToggleCount)
    Else
        cmb_DispToggleCount.text = "2" '�⺻�� 2ȸ
    End If
    
    
    
    'Lane Config
'    If LANE1_YN = "Y" Then
'        chk_UseYN(0).value = 1
'        If (Glo_Screen_No = 2) Then
'            Jung.ShapeCamera(0).FillColor = &HFF&
'        Else
'            FrmG4Mini.ShapeCamera(0).FillColor = &HFF&
'        End If
'    Else
'        If (Glo_Screen_No = 2) Then
'            Jung.ShapeCamera(0).FillColor = &H808080
'        Else
'            FrmG4Mini.ShapeCamera(0).FillColor = &H808080
'        End If
'    End If
    If (LANE1_YN = "Y") Then
        chk_UseYN(0).value = "1"
        Frame2(0).Enabled = True
    Else
        chk_UseYN(0).value = "0"
        Frame2(0).Enabled = False
    End If
    If (LANE1_Inout = "�Ա�") Then
        cmb_Inout(0).ListIndex = 0
    Else
        cmb_Inout(0).ListIndex = 1
    End If
    txt_GateName(0).text = LANE1_Name
    cmb_LPRMode(0).ListIndex = LANE1_LPRMode
    txt_LPRIP(0) = LANE1_LPRIP
    txt_LPRPort(0) = LANE1_LPRPort
    cmb_DeviceMode(0).ListIndex = LANE1_DeviceMode
    txt_DeviceIP(0).text = LANE1_DeviceIP
    cmb_DisplayMode(0).ListIndex = LANE1_DisplayMode
    txt_DispIP(0).text = LANE1_DispIP
    txt_DispPort(0).text = LANE1_DispPort
    txt_RelayPort(0).text = LANE1_RelayPort
    'cmb_DispComPort(0).ListIndex = (LANE1_DispComPort - 1)
    'cmb_RelayComPort(0).ListIndex = (LANE1_RelayComPort - 1)
    txt_Disp1(0) = LANE1_Disp1Msg
    txt_Disp2(0) = LANE1_Disp2Msg
    cmb_Disp1(0).ListIndex = LANE1_Disp1Color
    cmb_Disp2(0).ListIndex = LANE1_Disp2Color
    cmb_DispShiftSpeed(0).ListIndex = LANE1_DispSpeed
    If (Glo_LANE_DISP_NML_SHIFT(0) = enumDISP_NML_SHIFT.eSTOP) Then
        cmd_NmlShift(0).Caption = "����"
        cmb_DispShiftSpeed(0).Visible = False
    ElseIf (Glo_LANE_DISP_NML_SHIFT(0) = enumDISP_NML_SHIFT.eSHIFT) Then
        cmd_NmlShift(0).Caption = "�̵�"
        cmb_DispShiftSpeed(0).Visible = True
    End If
    
'    If LANE2_YN = "Y" Then
'        chk_UseYN(1).value = 1
'        If (Glo_Screen_No = 2) Then
'            Jung.ShapeCamera(1).FillColor = &HFF&
'        Else
'            FrmG4Mini.ShapeCamera(1).FillColor = &HFF&
'        End If
'    Else
'        If (Glo_Screen_No = 2) Then
'            Jung.ShapeCamera(1).FillColor = &H808080
'        Else
'            FrmG4Mini.ShapeCamera(1).FillColor = &H808080
'        End If
'    End If
    If (LANE2_YN = "Y") Then
        chk_UseYN(1).value = "1"
        Frame2(1).Enabled = True
    Else
        chk_UseYN(1).value = "0"
        Frame2(1).Enabled = False
    End If
    If (LANE2_Inout = "�Ա�") Then
        cmb_Inout(1).ListIndex = 0
    Else
        cmb_Inout(1).ListIndex = 1
    End If
    txt_GateName(1).text = LANE2_Name
    cmb_LPRMode(1).ListIndex = LANE2_LPRMode
    txt_LPRIP(1) = LANE2_LPRIP
    txt_LPRPort(1) = LANE2_LPRPort
    cmb_DeviceMode(1).ListIndex = LANE2_DeviceMode
    txt_DeviceIP(1).text = LANE2_DeviceIP
    cmb_DisplayMode(1).ListIndex = LANE2_DisplayMode
    txt_DispIP(1).text = LANE2_DispIP
    txt_DispPort(1).text = LANE2_DispPort
    txt_RelayPort(1).text = LANE2_RelayPort
    'cmb_DispComPort(1).ListIndex = (LANE2_DispComPort - 1)
    'cmb_RelayComPort(1).ListIndex = (LANE2_RelayComPort - 1)
    txt_Disp1(1) = LANE2_Disp1Msg
    txt_Disp2(1) = LANE2_Disp2Msg
    cmb_Disp1(1).ListIndex = LANE2_Disp1Color
    cmb_Disp2(1).ListIndex = LANE2_Disp2Color
    cmb_DispShiftSpeed(1).ListIndex = LANE2_DispSpeed
    If (Glo_LANE_DISP_NML_SHIFT(1) = enumDISP_NML_SHIFT.eSTOP) Then
        cmd_NmlShift(1).Caption = "����"
        cmb_DispShiftSpeed(1).Visible = False
    ElseIf (Glo_LANE_DISP_NML_SHIFT(1) = enumDISP_NML_SHIFT.eSHIFT) Then
        cmd_NmlShift(1).Caption = "�̵�"
        cmb_DispShiftSpeed(1).Visible = True
    End If
    
    
'    If LANE3_YN = "Y" Then
'        chk_UseYN(2).value = 1
'        If (Glo_Screen_No = 2) Then
'            Jung.ShapeCamera(2).FillColor = &HFF&
'        Else
'            FrmG4Mini.ShapeCamera(2).FillColor = &HFF&
'        End If
'    Else
'        If (Glo_Screen_No = 2) Then
'            Jung.ShapeCamera(2).FillColor = &H808080
'        Else
'            FrmG4Mini.ShapeCamera(2).FillColor = &H808080
'        End If
'    End If
    If (LANE3_YN = "Y") Then
        chk_UseYN(2).value = "1"
        Frame2(2).Enabled = True
    Else
        chk_UseYN(2).value = "0"
        Frame2(2).Enabled = False
    End If
    If (LANE3_Inout = "�Ա�") Then
        cmb_Inout(2).ListIndex = 0
    Else
        cmb_Inout(2).ListIndex = 1
    End If
    txt_GateName(2).text = LANE3_Name
    cmb_LPRMode(2).ListIndex = LANE3_LPRMode
    txt_LPRIP(2) = LANE3_LPRIP
    txt_LPRPort(2) = LANE3_LPRPort
    cmb_DeviceMode(2).ListIndex = LANE3_DeviceMode
    txt_DeviceIP(2).text = LANE3_DeviceIP
    cmb_DisplayMode(2).ListIndex = LANE3_DisplayMode
    txt_DispIP(2).text = LANE3_DispIP
    txt_DispPort(2).text = LANE3_DispPort
    txt_RelayPort(2).text = LANE3_RelayPort
    'cmb_DispComPort(2).ListIndex = (LANE3_DispComPort - 1)
    'cmb_RelayComPort(2).ListIndex = (LANE3_RelayComPort - 1)
    txt_Disp1(2) = LANE3_Disp1Msg
    txt_Disp2(2) = LANE3_Disp2Msg
    cmb_Disp1(2).ListIndex = LANE3_Disp1Color
    cmb_Disp2(2).ListIndex = LANE3_Disp2Color
    cmb_DispShiftSpeed(2).ListIndex = LANE3_DispSpeed
    If (Glo_LANE_DISP_NML_SHIFT(2) = enumDISP_NML_SHIFT.eSTOP) Then
        cmd_NmlShift(2).Caption = "����"
        cmb_DispShiftSpeed(2).Visible = False
    ElseIf (Glo_LANE_DISP_NML_SHIFT(2) = enumDISP_NML_SHIFT.eSHIFT) Then
        cmd_NmlShift(2).Caption = "�̵�"
        cmb_DispShiftSpeed(2).Visible = True
    End If
    
    
'    If LANE4_YN = "Y" Then
'        chk_UseYN(3).value = 1
'        If (Glo_Screen_No = 2) Then
'            Jung.ShapeCamera(3).FillColor = &HFF&
'        Else
'            FrmG4Mini.ShapeCamera(3).FillColor = &HFF&
'        End If
'    Else
'        If (Glo_Screen_No = 2) Then
'            Jung.ShapeCamera(3).FillColor = &H808080
'        Else
'            FrmG4Mini.ShapeCamera(3).FillColor = &H808080
'        End If
'    End If
    If (LANE4_YN = "Y") Then
        chk_UseYN(3).value = "1"
        Frame2(3).Enabled = True
    Else
        chk_UseYN(3).value = "0"
        Frame2(3).Enabled = False
    End If
    If (LANE4_Inout = "�Ա�") Then
        cmb_Inout(3).ListIndex = 0
    Else
        cmb_Inout(3).ListIndex = 1
    End If
    txt_GateName(3).text = LANE4_Name
    cmb_LPRMode(3).ListIndex = LANE4_LPRMode
    txt_LPRIP(3) = LANE4_LPRIP
    txt_LPRPort(3) = LANE4_LPRPort
    cmb_DeviceMode(3).ListIndex = LANE4_DeviceMode
    txt_DeviceIP(3).text = LANE4_DeviceIP
    cmb_DisplayMode(3).ListIndex = LANE4_DisplayMode
    txt_DispIP(3).text = LANE4_DispIP
    txt_DispPort(3).text = LANE4_DispPort
    txt_RelayPort(3).text = LANE4_RelayPort
    'cmb_DispComPort(3).ListIndex = (LANE4_DispComPort - 1)
    'cmb_RelayComPort(3).ListIndex = (LANE4_RelayComPort - 1)
    txt_Disp1(3) = LANE4_Disp1Msg
    txt_Disp2(3) = LANE4_Disp2Msg
    cmb_Disp1(3).ListIndex = LANE4_Disp1Color
    cmb_Disp2(3).ListIndex = LANE4_Disp2Color
    cmb_DispShiftSpeed(3).ListIndex = LANE4_DispSpeed
    If (Glo_LANE_DISP_NML_SHIFT(3) = enumDISP_NML_SHIFT.eSTOP) Then
        cmd_NmlShift(3).Caption = "����"
        cmb_DispShiftSpeed(3).Visible = False
    ElseIf (Glo_LANE_DISP_NML_SHIFT(3) = enumDISP_NML_SHIFT.eSHIFT) Then
        cmd_NmlShift(3).Caption = "�̵�"
        cmb_DispShiftSpeed(3).Visible = True
    End If
    
    
    If (LANE5_YN = "Y") Then
        chk_UseYN(4).value = "1"
        Frame2(4).Enabled = True
    Else
        chk_UseYN(4).value = "0"
        Frame2(4).Enabled = False
    End If
    If (LANE5_Inout = "�Ա�") Then
        cmb_Inout(4).ListIndex = 0
    Else
        cmb_Inout(4).ListIndex = 1
    End If
    txt_GateName(4).text = LANE5_Name
    cmb_LPRMode(4).ListIndex = LANE5_LPRMode
    txt_LPRIP(4) = LANE5_LPRIP
    txt_LPRPort(4) = LANE5_LPRPort
    cmb_DeviceMode(4).ListIndex = LANE5_DeviceMode
    txt_DeviceIP(4).text = LANE5_DeviceIP
    cmb_DisplayMode(4).ListIndex = LANE5_DisplayMode
    txt_DispIP(4).text = LANE5_DispIP
    txt_DispPort(4).text = LANE5_DispPort
    txt_RelayPort(4).text = LANE5_RelayPort
    txt_Disp1(4) = LANE5_Disp1Msg
    txt_Disp2(4) = LANE5_Disp2Msg
    cmb_Disp1(4).ListIndex = LANE5_Disp1Color
    cmb_Disp2(4).ListIndex = LANE5_Disp2Color
    cmb_DispShiftSpeed(4).ListIndex = LANE5_DispSpeed
    If (Glo_LANE_DISP_NML_SHIFT(4) = enumDISP_NML_SHIFT.eSTOP) Then
        cmd_NmlShift(4).Caption = "����"
        cmb_DispShiftSpeed(4).Visible = False
    ElseIf (Glo_LANE_DISP_NML_SHIFT(4) = enumDISP_NML_SHIFT.eSHIFT) Then
        cmd_NmlShift(4).Caption = "�̵�"
        cmb_DispShiftSpeed(4).Visible = True
    End If
    
 
    If LANE6_YN = "Y" Then
        chk_UseYN(5).value = 1
        Frame2(5).Enabled = True
    Else
        chk_UseYN(5).value = 0
        Frame2(5).Enabled = False
    End If
    If (LANE6_Inout = "�Ա�") Then
        cmb_Inout(5).ListIndex = 0
    Else
        cmb_Inout(5).ListIndex = 1
    End If
    txt_GateName(5).text = LANE6_Name
    cmb_LPRMode(5).ListIndex = LANE6_LPRMode
    txt_LPRIP(5) = LANE6_LPRIP
    txt_LPRPort(5) = LANE6_LPRPort
    cmb_DeviceMode(5).ListIndex = LANE6_DeviceMode
    txt_DeviceIP(5).text = LANE6_DeviceIP
    cmb_DisplayMode(5).ListIndex = LANE6_DisplayMode
    txt_DispIP(5).text = LANE6_DispIP
    txt_DispPort(5).text = LANE6_DispPort
    txt_RelayPort(5).text = LANE6_RelayPort
    txt_Disp1(5) = LANE6_Disp1Msg
    txt_Disp2(5) = LANE6_Disp2Msg
    cmb_Disp1(5).ListIndex = LANE6_Disp1Color
    cmb_Disp2(5).ListIndex = LANE6_Disp2Color
    cmb_DispShiftSpeed(5).ListIndex = LANE6_DispSpeed
    If (Glo_LANE_DISP_NML_SHIFT(5) = enumDISP_NML_SHIFT.eSTOP) Then
        cmd_NmlShift(5).Caption = "����"
        cmb_DispShiftSpeed(5).Visible = False
    ElseIf (Glo_LANE_DISP_NML_SHIFT(5) = enumDISP_NML_SHIFT.eSHIFT) Then
        cmd_NmlShift(5).Caption = "�̵�"
        cmb_DispShiftSpeed(5).Visible = True
    End If
    
    
    Cmb_Window.Clear
    Cmb_Window.AddItem "����ȭ��"
    Cmb_Window.AddItem "2����ȭ��"
    Cmb_Window.AddItem "4����ȭ��"
    Cmb_Window.AddItem "6����ȭ��"
    
    
    
    
    Glo_Screen1 = Get_Ini("System Config", "LANE1_ȭ����ġ", "1")
    Glo_Screen2 = Get_Ini("System Config", "LANE2_ȭ����ġ", "2")
    Glo_Screen3 = Get_Ini("System Config", "LANE3_ȭ����ġ", "3")
    Glo_Screen4 = Get_Ini("System Config", "LANE4_ȭ����ġ", "4")
    Glo_Screen5 = Get_Ini("System Config", "LANE5_ȭ����ġ", "5")
    Glo_Screen6 = Get_Ini("System Config", "LANE6_ȭ����ġ", "6")
    
    
    'If (cmb_Board.text = "�����") Then
        'Frame4.Caption = "����̽� ��Ź��"
        Frame8.Caption = "�����ġ ����"
        
        '�����ġ
        Cmb_Display.Clear
        Cmb_Display.AddItem "������"
        'Cmb_Display.AddItem "������(Full Color)"
        Cmb_Display.AddItem "FND"
        Cmb_Display.AddItem "������(Ǯ�÷�)"
        Cmb_Display.AddItem "������(Ǯ�÷�)_FW7"
        
        If (Glo_Display = "������") Then
            Cmb_Display.ListIndex = 0
        ElseIf (Glo_Display = "FND") Then
            Cmb_Display.ListIndex = 1
        ElseIf (Glo_Display = "������(Ǯ�÷�)") Then
            Cmb_Display.ListIndex = 2
        ElseIf (Glo_Display = "������(Ǯ�÷�)_FW7") Then
            Cmb_Display.ListIndex = 3
        Else
            Cmb_Display.ListIndex = 3
        End If
    'End If
    
    '������ ��¹���
    Cmb_Display_Direct.Clear
    Cmb_Display_Direct.AddItem "����"
    Cmb_Display_Direct.AddItem "����"
    If (Glo_Display_Direct = "����") Then
        Cmb_Display_Direct.ListIndex = 0
    Else
        Cmb_Display_Direct.ListIndex = 1
    End If



    Call SetCommunication



    cmb_PrintPort(0).text = Glo_Guest_Print_Port(0)
    cmb_PrintPort(1).text = Glo_Guest_Print_Port(1)
    cmb_PrintPort(2).text = Glo_Guest_Print_Port(2)
    cmb_PrintPort(3).text = Glo_Guest_Print_Port(3)
    cmb_PrintPort(4).text = Glo_Guest_Print_Port(4)
    cmb_PrintPort(5).text = Glo_Guest_Print_Port(5)
    
    cmb_PrintModel(0).text = Glo_Guest_Print_Model(0)
    cmb_PrintModel(1).text = Glo_Guest_Print_Model(1)
    cmb_PrintModel(2).text = Glo_Guest_Print_Model(2)
    cmb_PrintModel(3).text = Glo_Guest_Print_Model(3)
    cmb_PrintModel(4).text = Glo_Guest_Print_Model(4)
    cmb_PrintModel(5).text = Glo_Guest_Print_Model(5)
    
    Call Print_Port_Init(0, Glo_GUEST_LANE1_YN, Glo_Guest_Print_Model(0), Glo_Guest_Print_Port(0))
    Call Print_Port_Init(1, Glo_GUEST_LANE2_YN, Glo_Guest_Print_Model(1), Glo_Guest_Print_Port(1))
    Call Print_Port_Init(2, Glo_GUEST_LANE3_YN, Glo_Guest_Print_Model(2), Glo_Guest_Print_Port(2))
    Call Print_Port_Init(3, Glo_GUEST_LANE4_YN, Glo_Guest_Print_Model(3), Glo_Guest_Print_Port(3))
    Call Print_Port_Init(4, Glo_GUEST_LANE5_YN, Glo_Guest_Print_Model(4), Glo_Guest_Print_Port(4))
    Call Print_Port_Init(5, Glo_GUEST_LANE6_YN, Glo_Guest_Print_Model(5), Glo_Guest_Print_Port(5))
    
'    CmbScreen1.Clear
'    CmbScreen2.Clear
'    CmbScreen3.Clear
'    CmbScreen4.Clear
    
    
    Select Case Glo_Screen_No
           Case 6
                Cmb_Window.ListIndex = 3 '0:����ȭ��, 1:2ȭ��, 2:4ȭ��, 3:6ȭ��
                
                For i = 0 To Glo_Screen_No - 1
                    CmbScreen(i).Clear
                    CmbScreen(i).Enabled = True
                    CmbScreen(i).AddItem "�⺻��ġ"
                    CmbScreen(i).AddItem "1����ġ"
                    CmbScreen(i).AddItem "2����ġ"
                    CmbScreen(i).AddItem "3����ġ"
                    CmbScreen(i).AddItem "4����ġ"
                    CmbScreen(i).AddItem "5����ġ"
                    CmbScreen(i).AddItem "6����ġ"
                Next
                
                bScrNoChk = True
                For i = 0 To Glo_Screen_No - 1
                    If (Glo_Screen1 < 1 Or Glo_Screen1 > Glo_Screen_No Or Glo_Screen2 < 1 Or Glo_Screen2 > Glo_Screen_No Or Glo_Screen3 < 1 Or Glo_Screen3 > Glo_Screen_No Or Glo_Screen4 < 1 Or Glo_Screen4 > Glo_Screen_No Or Glo_Screen5 < 1 Or Glo_Screen5 > Glo_Screen_No Or Glo_Screen6 < 1 Or Glo_Screen6 > Glo_Screen_No) Then
                        bScrNoChk = False
                        Exit For
                    End If
                Next
                If (bScrNoChk = True) Then
                    CmbScreen(0).ListIndex = Glo_Screen1
                    CmbScreen(1).ListIndex = Glo_Screen2
                    CmbScreen(2).ListIndex = Glo_Screen3
                    CmbScreen(3).ListIndex = Glo_Screen4
                    CmbScreen(4).ListIndex = Glo_Screen5
                    CmbScreen(5).ListIndex = Glo_Screen6
                Else
                    CmbScreen(0).ListIndex = 1
                    CmbScreen(1).ListIndex = 2
                    CmbScreen(2).ListIndex = 3
                    CmbScreen(3).ListIndex = 4
                    CmbScreen(4).ListIndex = 5
                    CmbScreen(5).ListIndex = 6
                End If
                
'                If (Glo_ApsYN = "Y") Then
'                    FrmG6_23.cmd_menu(8).Visible = True
'                    'FrmG6_23.cmd_menu(8).Enabled = True
'                Else
'                    FrmG6_23.cmd_menu(8).Visible = False
'                    'FrmG6_23.cmd_menu(8).Enabled = False
'                End If
                

                FrmG6_23.Frame1(0).Left = 7 + Int((CmbScreen(0).ListIndex - 1) Mod 3) * 636
                FrmG6_23.Frame1(0).Top = 70 + Int((CmbScreen(0).ListIndex - 1) / 3) * 481
                FrmG6_23.Frame1(1).Left = 7 + Int((CmbScreen(1).ListIndex - 1) Mod 3) * 636
                FrmG6_23.Frame1(1).Top = 70 + Int((CmbScreen(1).ListIndex - 1) / 3) * 481
                FrmG6_23.Frame1(2).Left = 7 + Int((CmbScreen(2).ListIndex - 1) Mod 3) * 636
                FrmG6_23.Frame1(2).Top = 70 + Int((CmbScreen(2).ListIndex - 1) / 3) * 481
                FrmG6_23.Frame1(3).Left = 7 + Int((CmbScreen(3).ListIndex - 1) Mod 3) * 636
                FrmG6_23.Frame1(3).Top = 70 + Int((CmbScreen(3).ListIndex - 1) / 3) * 481
                FrmG6_23.Frame1(4).Left = 7 + Int((CmbScreen(4).ListIndex - 1) Mod 3) * 636
                FrmG6_23.Frame1(4).Top = 70 + Int((CmbScreen(4).ListIndex - 1) / 3) * 481
                FrmG6_23.Frame1(5).Left = 7 + Int((CmbScreen(5).ListIndex - 1) Mod 3) * 636
                FrmG6_23.Frame1(5).Top = 70 + Int((CmbScreen(5).ListIndex - 1) / 3) * 481
                
                
           Case 1
                Cmb_Window.ListIndex = 0
                
                For i = 0 To MAX_LANE_COUNT - 1
                    CmbScreen(i).Enabled = False
                Next
                For i = 0 To Glo_Screen_No - 1
                    CmbScreen(i).Enabled = True
                Next
                
'                If (Glo_ApsYN = "Y") Then
'                    FrmG1.Lblbutton(7).Visible = True
'                    FrmG1.Imgbutton(7).Visible = True
'                    FrmG1.Lblbutton(8).Visible = True
'                    FrmG1.Imgbutton(8).Visible = True
'                Else
'                    FrmG1.Lblbutton(7).Visible = False
'                    FrmG1.Imgbutton(7).Visible = False
'                    FrmG1.Lblbutton(8).Visible = False
'                    FrmG1.Imgbutton(8).Visible = False
'                End If
                
           Case 2
                Cmb_Window.ListIndex = 1
                
                For i = 0 To MAX_LANE_COUNT - 1
                    CmbScreen(i).Enabled = False
                Next
                For i = 0 To Glo_Screen_No - 1
                    CmbScreen(i).Clear
                    CmbScreen(i).Enabled = True
'                    CmbScreen(i).AddItem "�⺻��ġ"
'                    CmbScreen(i).AddItem "��ġ����"
                    CmbScreen(i).AddItem "�⺻��ġ"
                    CmbScreen(i).AddItem "1����ġ"
                    CmbScreen(i).AddItem "2����ġ"
                Next
                
                bScrNoChk = True
                For i = 0 To Glo_Screen_No - 1
                    If (Glo_Screen1 < 1 Or Glo_Screen1 > Glo_Screen_No Or Glo_Screen2 < 1 Or Glo_Screen2 > Glo_Screen_No) Then
                        bScrNoChk = False
                        Exit For
                    End If
                Next
                If (bScrNoChk = True) Then
                    CmbScreen(0).ListIndex = Glo_Screen1
                    CmbScreen(1).ListIndex = Glo_Screen2
                Else
                    CmbScreen(0).ListIndex = 1
                    CmbScreen(1).ListIndex = 2
                End If

                If (Glo_Screen1 = 1) Then
                    Jung.Frame1(0).Left = 120
                    Jung.Frame1(0).Top = 2070
                    Jung.Frame1(1).Left = 13065
                    Jung.Frame1(1).Top = 2070
                Else
                    Jung.Frame1(0).Left = 13065
                    Jung.Frame1(0).Top = 2070
                    Jung.Frame1(1).Left = 120
                    Jung.Frame1(1).Top = 2070
                End If
                
'                If (Glo_ApsYN = "Y") Then
'                    Jung.Lblbutton(7).Visible = True
'                    Jung.Imgbutton(7).Visible = True
'                    Jung.Lblbutton(8).Visible = True
'                    Jung.Imgbutton(8).Visible = True
'                Else
'                    Jung.Lblbutton(7).Visible = False
'                    Jung.Imgbutton(7).Visible = False
'                    Jung.Lblbutton(8).Visible = False
'                    Jung.Imgbutton(8).Visible = False
'                End If
                
           Case 4
           
                Dim itop As Integer
                Dim left1 As Integer
                Dim left2 As Integer
                Dim left3 As Integer
                Dim left4 As Integer

                Cmb_Window.ListIndex = 2
                
                For i = 0 To MAX_LANE_COUNT - 1
                    CmbScreen(i).Enabled = False
                Next
                For i = 0 To Glo_Screen_No - 1
                    CmbScreen(i).Clear
                    CmbScreen(i).Enabled = True
                    CmbScreen(i).AddItem "�⺻��ġ"
                    CmbScreen(i).AddItem "1����ġ"
                    CmbScreen(i).AddItem "2����ġ"
                    CmbScreen(i).AddItem "3����ġ"
                    CmbScreen(i).AddItem "4����ġ"
                Next

                bScrNoChk = True
                For i = 0 To Glo_Screen_No - 1
                    If (Glo_Screen1 < 1 Or Glo_Screen1 > Glo_Screen_No Or Glo_Screen2 < 1 Or Glo_Screen2 > Glo_Screen_No Or Glo_Screen3 < 1 Or Glo_Screen3 > Glo_Screen_No Or Glo_Screen4 < 1 Or Glo_Screen4 > Glo_Screen_No) Then
                        bScrNoChk = False
                        Exit For
                    End If
                Next
                If (bScrNoChk = True) Then
                    CmbScreen(0).ListIndex = Glo_Screen1
                    CmbScreen(1).ListIndex = Glo_Screen2
                    CmbScreen(2).ListIndex = Glo_Screen3
                    CmbScreen(3).ListIndex = Glo_Screen4
                Else
                    CmbScreen(0).ListIndex = 1
                    CmbScreen(1).ListIndex = 2
                    CmbScreen(2).ListIndex = 3
                    CmbScreen(3).ListIndex = 4
                End If
    

                FrmG4Mini.Frame1(0).Left = 3 + Int((CmbScreen(0).ListIndex - 1) Mod 4) * 319
                FrmG4Mini.Frame1(0).Top = 131
                FrmG4Mini.Frame1(1).Left = 3 + Int((CmbScreen(1).ListIndex - 1) Mod 4) * 319
                FrmG4Mini.Frame1(1).Top = 131
                FrmG4Mini.Frame1(2).Left = 3 + Int((CmbScreen(2).ListIndex - 1) Mod 4) * 319
                FrmG4Mini.Frame1(2).Top = 131
                FrmG4Mini.Frame1(3).Left = 3 + Int((CmbScreen(3).ListIndex - 1) Mod 4) * 319
                FrmG4Mini.Frame1(3).Top = 131
                
'                If (Glo_ApsYN = "Y") Then
'                    FrmG4Mini.Lblbutton(7).Visible = True
'                    FrmG4Mini.Imgbutton(7).Visible = True
'                    FrmG4Mini.Lblbutton(8).Visible = True
'                    FrmG4Mini.Imgbutton(8).Visible = True
'                Else
'                    FrmG4Mini.Lblbutton(7).Visible = False
'                    FrmG4Mini.Imgbutton(7).Visible = False
'                    FrmG4Mini.Lblbutton(8).Visible = False
'                    FrmG4Mini.Imgbutton(8).Visible = False
'                End If
    End Select
    
    
    
    
    Timer1.Enabled = True
    DBTimer.Enabled = True
    DB_Connect_Timer = False


'''    '������ �ӵ� �ε�
'''    Dim rs As Recordset
'''    Set rs = New ADODB.Recordset
'''    rs.Open "SELECT * FROM tb_config", adoConn
'''    Do While Not (rs.EOF)
'''        Select Case rs!Title
'''            Case "LANE1_DISP_SPEED"
'''                cmb_ShiftSpeed(0).ListIndex = rs!CONTENT
'''            Case "LANE2_DISP_SPEED"
'''                cmb_ShiftSpeed(1).ListIndex = rs!CONTENT
'''            Case "LANE3_DISP_SPEED"
'''                cmb_ShiftSpeed(2).ListIndex = rs!CONTENT
'''            Case "LANE4_DISP_SPEED"
'''                cmb_ShiftSpeed(3).ListIndex = rs!CONTENT
'''            Case "LANE5_DISP_SPEED"
'''                cmb_ShiftSpeed(4).ListIndex = rs!CONTENT
'''            Case "LANE6_DISP_SPEED"
'''                cmb_ShiftSpeed(5).ListIndex = rs!CONTENT
'''        End Select
'''        rs.MoveNext
'''    Loop
'''    Set rs = Nothing
    
    
    '����
    Call ParkFull_Set
    
    '������
    Call ParkFullLight_Set
    
    'Call SetCommunication
    
    Call ShowTitlebarSiteCode
    
    '���ܱ�ݱ� UI
    Call ShowGateClose
    
    '�湮�������� �����ε�
    Call LoadGuestReg_YN
    
    '������ ���� �ε�
    Call LoadWebDC_YN
    
    
    '����Ͼ˸� ��ư
    Call LoadMobileAlarm_YN
    
Exit Sub

Err_Proc:
    MsgBox ("[FormLoad_Proc]  " & Err.Description)
    Call DataLogger(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    [TCP Server Load Proc]  " & Err.Description)

End Sub

Public Sub LoadMobileAlarm_YN()
    Dim bQryResult As Boolean
    Dim sQry As String
    Dim rsMAlarm As ADODB.Recordset
    
    'Glo_MobileAlarm_YN = "N"
    
On Error GoTo Err_p
    
    sQry = "SELECT * FROM tb_config where TITLE = 'MOBILE' AND NAME = 'ALARM' "
    Set rsMAlarm = New ADODB.Recordset
    bQryResult = DataBaseQuery(rsMAlarm, adoConn, sQry, NWERR_GATE_STAY)
    If (bQryResult = False) Then
        DataLogger ("[LoadMobileAlarm]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
        Exit Sub
    End If
    
    If Not (rsMAlarm.EOF) Then
        Glo_MobileAlarm_YN = "" & rsMAlarm!Content
    Else
        Glo_MobileAlarm_YN = "N"
    End If
    
    Set rsMAlarm = Nothing
    
    
    If (Glo_MobileAlarm_YN = "Y") Then
        Command10.Enabled = True
        Command10.Visible = True
    Else
        Command10.Enabled = False
        Command10.Visible = False
    End If
    
    
    Exit Sub
    
Err_p:
    Set rsMAlarm = Nothing
    Command10.Enabled = False
    Command10.Visible = False
    Call DataLogger(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    [TCP Server Load LoadWebDC]  " & Err.Description)
End Sub
Public Sub LoadWebDC_YN()
    'Glo_WebDC_YN = "N"
    Dim bQryResult As Boolean
    Dim sQry As String
    Dim rsWebDC As ADODB.Recordset
    
    Glo_WebDC_YN = "N"
    
On Error GoTo Err_p
    
    sQry = "SELECT * FROM tb_config where NAME = 'WebDC' "
    Set rsWebDC = New ADODB.Recordset
    'rsWebDC.Open Qry, adoConn
    bQryResult = DataBaseQuery(rsWebDC, adoConn, sQry, NWERR_GATE_STAY)
    If (bQryResult = False) Then
        DataLogger ("[LoadWebDC]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
        Exit Sub
    End If
    
    If Not (rsWebDC.EOF) Then
        Glo_WebDC_YN = "" & rsWebDC!Content
    End If
    
    Set rsWebDC = Nothing
    
    Exit Sub
Err_p:
    Set rsWebDC = Nothing
    Call DataLogger(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    [TCP Server Load LoadWebDC]  " & Err.Description)
End Sub

'�湮���� ���� �ε�
Public Sub LoadGuestReg_YN()
    
    Dim bQryResult As Boolean
    Dim sGuestQry As String
    Dim rsGuestReg As ADODB.Recordset
    
    Glo_GuestReg_YN = "N"
    
On Error GoTo Err_p
    
    sGuestQry = "SELECT * FROM tb_config where NAME = 'GuestCarReg' "
    Set rsGuestReg = New ADODB.Recordset
    'rsGuestReg.Open Qry, adoConn
    bQryResult = DataBaseQuery(rsGuestReg, adoConn, sGuestQry, NWERR_GATE_STAY)
    If (bQryResult = False) Then
        DataLogger ("[LoadGuestReg]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
        Exit Sub
    End If
    
    If Not (rsGuestReg.EOF) Then
        Glo_GuestReg_YN = "" & rsGuestReg!Content
    End If
    
    Set rsGuestReg = Nothing
    
    Exit Sub
Err_p:
    Set rsGuestReg = Nothing
    Call DataLogger(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    [TCP Server Load LoadGuestReg]  " & Err.Description)
End Sub

Public Sub ShowGateClose()
    
   Select Case Glo_Screen_No
        Case 6
        Case 4
            If (Glo_Lane1_GateClose_YN = "Y") Then
                FrmG4Mini.cmd_GateClose(0).Visible = True
                FrmG4Mini.cmd_GateClose(0).Enabled = True
            Else
                FrmG4Mini.cmd_GateClose(0).Visible = False
                FrmG4Mini.cmd_GateClose(0).Enabled = False
            End If
            If (Glo_Lane2_GateClose_YN = "Y") Then
                FrmG4Mini.cmd_GateClose(1).Visible = True
                FrmG4Mini.cmd_GateClose(1).Enabled = True
            Else
                FrmG4Mini.cmd_GateClose(1).Visible = False
                FrmG4Mini.cmd_GateClose(1).Enabled = False
            End If
            If (Glo_Lane3_GateClose_YN = "Y") Then
                FrmG4Mini.cmd_GateClose(2).Visible = True
                FrmG4Mini.cmd_GateClose(2).Enabled = True
            Else
                FrmG4Mini.cmd_GateClose(2).Visible = False
                FrmG4Mini.cmd_GateClose(2).Enabled = False
            End If
            If (Glo_Lane4_GateClose_YN = "Y") Then
                FrmG4Mini.cmd_GateClose(3).Visible = True
                FrmG4Mini.cmd_GateClose(3).Enabled = True
            Else
                FrmG4Mini.cmd_GateClose(3).Visible = False
                FrmG4Mini.cmd_GateClose(3).Enabled = False
            End If
        Case 2
            If (Glo_Lane1_GateClose_YN = "Y") Then
                Jung.cmd_GateClose(0).Visible = True
                Jung.cmd_GateClose(0).Enabled = True
            Else
                Jung.cmd_GateClose(0).Visible = False
                Jung.cmd_GateClose(0).Enabled = False
            End If
            If (Glo_Lane2_GateClose_YN = "Y") Then
                Jung.cmd_GateClose(1).Visible = True
                Jung.cmd_GateClose(1).Enabled = True
            Else
                Jung.cmd_GateClose(1).Visible = False
                Jung.cmd_GateClose(1).Enabled = False
            End If
        Case 1
            If (Glo_Lane1_GateClose_YN = "Y") Then
                FrmG1.cmd_GateClose(0).Visible = True
                FrmG1.cmd_GateClose(0).Enabled = True
            Else
                FrmG1.cmd_GateClose(0).Visible = False
                FrmG1.cmd_GateClose(0).Enabled = False
            End If
            
    End Select

End Sub

'============================================================================================================
' ���ȯ�� ����
'============================================================================================================
Public Sub SetCommunication()

'On Error GoTo Err_Proc

    Call CloseAllSock   '��� ���� �ݱ�
    
    Call Server.StartServer(Server_Port, Server.ServerIP) 'TCP ���� ����
'''    Call Server_WebDC.StartServer(Server_WEBDCPort, "0.0.0.0") '���������� ���μ������Ϲ���(��DC Ŭ���̾�Ʈ�κ��� ���ܱ� ���� ��� ���ſ�)


            '���ܱ����� Ŭ���̾�Ʈ ���� ��Ʈ����
            If (Glo_GateAgent_YN = "Y") Then
                LANE1_RelayPort = Glo_GATE_AGENT1_PORT: LANE2_RelayPort = Glo_GATE_AGENT2_PORT: LANE3_RelayPort = Glo_GATE_AGENT3_PORT:
                LANE4_RelayPort = Glo_GATE_AGENT4_PORT: LANE5_RelayPort = Glo_GATE_AGENT5_PORT: LANE6_RelayPort = Glo_GATE_AGENT6_PORT:
            Else
                LANE1_RelayPort = 1100: LANE2_RelayPort = 1100: LANE3_RelayPort = 1100: '����� ��Ʈ ���Ͽ� ���ܱ�����(ethernet)
                LANE4_RelayPort = 1100: LANE5_RelayPort = 1100: LANE6_RelayPort = 1100:
            End If
            
            
            '���������� Ŭ���̾�Ʈ ���� ��Ʈ����
            If (cmb_Board.text = "�����") Then
                LANE1_DispPort = 1000: LANE2_DispPort = 1000: LANE3_DispPort = 1000: '����� ��Ʈ ���Ͽ� ����������(ethernet)
                LANE4_DispPort = 1000: LANE5_DispPort = 1000: LANE6_DispPort = 1000:
            
            ElseIf (Glo_LPRBoard = "�ڵ��̳�") Then
                If (LANE1_DisplayMode = 0) Then 'TCP
                    LANE1_DispPort = 5000: LANE2_DispPort = 5000: LANE3_DispPort = 5000: '������ ���� ���� ��Ʈ(ethernet)
                    LANE4_DispPort = 5000: LANE5_DispPort = 5000: LANE6_DispPort = 5000:
                Else                            'UDP
                    LANE1_DispPort = 5108: LANE2_DispPort = 5108: LANE3_DispPort = 5108: '������ ���� ���� ��Ʈ(ethernet)
                    LANE4_DispPort = 5108: LANE5_DispPort = 5108: LANE6_DispPort = 5108:
                End If
            End If
        
        
            Dim lprPORT As Long
            Dim dispIP, deviceIP As String
            Dim dispPort, devicePORT As Long
            Dim i As Integer
            
            For i = 0 To MAX_LANE_COUNT - 1
                Select Case i
                    Case 0
                        lprPORT = LANE1_LPRPort:
                        deviceIP = LANE1_DeviceIP:  devicePORT = LANE1_RelayPort:   dispIP = LANE1_DispIP:  dispPort = LANE1_DispPort
                        LANE1_Handle = FindWindow(vbNullString, "Lane1"):   SendMess WM_HOST_HANDLE & gHW, LANE1_Handle
                    Case 1
                        lprPORT = LANE2_LPRPort
                        deviceIP = LANE2_DeviceIP:  devicePORT = LANE2_RelayPort:   dispIP = LANE2_DispIP:  dispPort = LANE2_DispPort
                        LANE2_Handle = FindWindow(vbNullString, "Lane2"):   SendMess WM_HOST_HANDLE & gHW, LANE2_Handle
                    Case 2
                        lprPORT = LANE3_LPRPort
                        deviceIP = LANE3_DeviceIP:  devicePORT = LANE3_RelayPort:   dispIP = LANE3_DispIP:  dispPort = LANE3_DispPort
                        LANE3_Handle = FindWindow(vbNullString, "Lane3"):   SendMess WM_HOST_HANDLE & gHW, LANE3_Handle
                    Case 3
                        lprPORT = LANE4_LPRPort
                        deviceIP = LANE4_DeviceIP:  devicePORT = LANE4_RelayPort:   dispIP = LANE4_DispIP:  dispPort = LANE4_DispPort
                        LANE4_Handle = FindWindow(vbNullString, "Lane4"):   SendMess WM_HOST_HANDLE & gHW, LANE4_Handle
                    Case 4
                        lprPORT = LANE5_LPRPort
                        deviceIP = LANE5_DeviceIP:  devicePORT = LANE5_RelayPort:   dispIP = LANE5_DispIP:  dispPort = LANE5_DispPort
                        LANE5_Handle = FindWindow(vbNullString, "Lane5"):   SendMess WM_HOST_HANDLE & gHW, LANE5_Handle
                    Case 5
                        lprPORT = LANE6_LPRPort
                        deviceIP = LANE6_DeviceIP:  devicePORT = LANE6_RelayPort:   dispIP = LANE6_DispIP:  dispPort = LANE6_DispPort
                        LANE6_Handle = FindWindow(vbNullString, "Lane6"):   SendMess WM_HOST_HANDLE & gHW, LANE6_Handle
                End Select
                
                '�߰�
                '����Ʈ������Ʈ�� ȣ��Ʈ�� ���ܱ� �Ƶ��̳���� �߰輭�� ����Ǹ�(C#) vb6�� ����̽� �߰����� �߰輭�� �����Ѵ�.
                'vb6���� ����̽��� ���� ���۽� ���ۼӵ�������������
                
                
                LPR_Send_sock(i).Protocol = sckUDPProtocol
                LPR1_sock(i).Protocol = sckUDPProtocol
                LPR1_sock(i).LocalPort = lprPORT '10101
                LPR1_sock(i).Bind
                
                
                If (cmb_Board.text = "�����") Then
                    Select Case LANE1_DeviceMode
                        Case "0"    'TCP
                            Disp1_sock(i).Protocol = sckTCPProtocol
                            Gate1_sock(i).Protocol = sckTCPProtocol
                        Case "1"    'UDP Only Send
                            Disp1_sock(i).Protocol = sckUDPProtocol
                            Disp1_sock(i).RemoteHost = dispIP
                            Disp1_sock(i).RemotePort = dispPort '1000
                            Gate1_sock(i).Protocol = sckUDPProtocol
                            Gate1_sock(i).RemoteHost = deviceIP
                            Gate1_sock(i).RemotePort = devicePORT '1100
                    End Select
                    

                ElseIf (cmb_Board.text = "�ڵ��̳�") Then
                    
                    '������:���� ������ ����(TCP, UDP)
                    Select Case LANE1_DisplayMode
                        Case "0"    'TCP
                            Disp1_sock(i).Protocol = sckTCPProtocol
                        Case "1"    'UDP Only Send
                            Disp1_sock(i).Protocol = sckUDPProtocol
                            Disp1_sock(i).RemoteHost = dispIP
                            Disp1_sock(i).RemotePort = dispPort
                    End Select
                    
                    '�ڵ��̳�:���ܱ� ����(TCP)
                    '����� VB -> �ڵ��̳� TCPó���ӵ����� ===> VB -> (UDP) -> C# middle ware -> (TCP) -> �ڵ��̳� ������� ó����
                    Select Case LANE1_DeviceMode
                        Case "0"    'TCP
                            Gate1_sock(i).Protocol = sckTCPProtocol
                        Case "1"    'UDP Only Send
                            Gate1_sock(i).Protocol = sckUDPProtocol
                            Gate1_sock(i).RemoteHost = deviceIP
                            Gate1_sock(i).RemotePort = devicePORT '1100
                    End Select
                End If
                
'''                Set Gate1_UniSock(i) = New UniSock 'TCP �ڵ��̳� ����̽�(���ܱ� ����)
'''
                '�ڵ��̳� ���� Ŀ�ǵ� UDP ���� ����
                Reset_sock(i).Protocol = sckUDPProtocol     '�ڵ��̳� Reset ó��
                Reset_sock(i).RemoteHost = deviceIP
                Reset_sock(i).RemotePort = devicePORT
            Next i
            
            '�ڵ��̳�κ��� ���� ����޼��� ó�� ����
            DeviceR_sock.Protocol = sckUDPProtocol
            DeviceR_sock.LocalPort = 3000
            DeviceR_sock.Bind
            
'''            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''        '    If (LANE2_YN = "Y") Then
'''                LPR_Send_sock(1).Protocol = sckUDPProtocol
'''                'LPR Engine
'''                LPR1_sock(1).Protocol = sckUDPProtocol
'''                LPR1_sock(1).LocalPort = LANE2_LPRPort
'''                LPR1_sock(1).Bind
'''                LANE2_Handle = FindWindow(vbNullString, "Lane2")
'''                SendMess WM_HOST_HANDLE & gHW, LANE2_Handle
'''
'''                'Wiznet Or COM Connection
'''                Select Case LANE1_DeviceMode
'''                    Case "0"    'TCP
'''                        Disp1_sock(1).Protocol = sckTCPProtocol
'''                        Gate1_sock(1).Protocol = sckTCPProtocol
'''                    Case "1"    'UDP Only Send
'''                        Disp1_sock(1).Protocol = sckUDPProtocol
'''                        Disp1_sock(1).RemoteHost = LANE2_DispIP
'''                        Disp1_sock(1).RemotePort = LANE2_DispPort
'''                        Gate1_sock(1).Protocol = sckUDPProtocol
'''                        Gate1_sock(1).RemoteHost = LANE2_DeviceIP
'''                        Gate1_sock(1).RemotePort = LANE2_RelayPort
'''                End Select
'''        '    End If
'''        '    If (LANE3_YN = "Y") Then
'''                LPR_Send_sock(2).Protocol = sckUDPProtocol
'''                'LPR Engine
'''                LPR1_sock(2).Protocol = sckUDPProtocol
'''                LPR1_sock(2).LocalPort = LANE3_LPRPort
'''                LPR1_sock(2).Bind
'''                LANE3_Handle = FindWindow(vbNullString, "Lane3")
'''                SendMess WM_HOST_HANDLE & gHW, LANE3_Handle
'''
'''
'''                'Wiznet Or COM Connection
'''                Select Case LANE1_DeviceMode
'''                    Case "0"    'TCP
'''                        Disp1_sock(2).Protocol = sckTCPProtocol
'''                        Gate1_sock(2).Protocol = sckTCPProtocol
'''                    Case "1"    'UDP Only Send
'''                        Disp1_sock(2).Protocol = sckUDPProtocol
'''                        Disp1_sock(2).RemoteHost = LANE3_DispIP
'''                        Disp1_sock(2).RemotePort = LANE3_DispPort
'''                        Gate1_sock(2).Protocol = sckUDPProtocol
'''                        Gate1_sock(2).RemoteHost = LANE3_DeviceIP
'''                        Gate1_sock(2).RemotePort = LANE3_RelayPort
'''                End Select
'''        '    End If
'''        '    If (LANE4_YN = "Y") Then
'''                LPR_Send_sock(3).Protocol = sckUDPProtocol
'''                'LPR Engine
'''                LPR1_sock(3).Protocol = sckUDPProtocol
'''                LPR1_sock(3).LocalPort = LANE4_LPRPort
'''                LPR1_sock(3).Bind
'''                LANE4_Handle = FindWindow(vbNullString, "Lane4")
'''                SendMess WM_HOST_HANDLE & gHW, LANE4_Handle
'''
'''                'Wiznet Or COM Connection
'''                Select Case LANE1_DeviceMode
'''                    Case "0"    'TCP
'''                        Disp1_sock(3).Protocol = sckTCPProtocol
'''                        Gate1_sock(3).Protocol = sckTCPProtocol
'''                    Case "1"    'UDP Only Send
'''                        Disp1_sock(3).Protocol = sckUDPProtocol
'''                        Disp1_sock(3).RemoteHost = LANE4_DispIP
'''                        Disp1_sock(3).RemotePort = LANE4_DispPort
'''                        Gate1_sock(3).Protocol = sckUDPProtocol
'''                        Gate1_sock(3).RemoteHost = LANE4_DeviceIP
'''                        Gate1_sock(3).RemotePort = LANE4_RelayPort
'''                End Select
'''        '    End If
'''        '    If (LANE5_YN = "Y") Then
'''                LPR_Send_sock(4).Protocol = sckUDPProtocol
'''                'LPR Engine
'''                LPR1_sock(4).Protocol = sckUDPProtocol
'''                LPR1_sock(4).LocalPort = LANE5_LPRPort
'''                LPR1_sock(4).Bind
'''                LANE5_Handle = FindWindow(vbNullString, "Lane5")
'''                SendMess WM_HOST_HANDLE & gHW, LANE5_Handle
'''
'''                'Wiznet Or COM Connection
'''                Select Case LANE1_DeviceMode
'''                    Case "0"    'TCP
'''                        Disp1_sock(4).Protocol = sckTCPProtocol
'''                        Gate1_sock(4).Protocol = sckTCPProtocol
'''                    Case "1"    'UDP Only Send
'''                        Disp1_sock(4).Protocol = sckUDPProtocol
'''                        Disp1_sock(4).RemoteHost = LANE5_DispIP
'''                        Disp1_sock(4).RemotePort = LANE5_DispPort
'''                        Gate1_sock(4).Protocol = sckUDPProtocol
'''                        Gate1_sock(4).RemoteHost = LANE5_DeviceIP
'''                        Gate1_sock(4).RemotePort = LANE5_RelayPort
'''                End Select
'''        '    End If
'''        '    If (LANE6_YN = "Y") Then
'''                LPR_Send_sock(5).Protocol = sckUDPProtocol
'''                'LPR Engine
'''                LPR1_sock(5).Protocol = sckUDPProtocol
'''                LPR1_sock(5).LocalPort = LANE6_LPRPort
'''                LPR1_sock(5).Bind
'''                LANE6_Handle = FindWindow(vbNullString, "Lane6")
'''                SendMess WM_HOST_HANDLE & gHW, LANE6_Handle
'''
'''                'Wiznet Or COM Connection
'''                Select Case LANE1_DeviceMode
'''                    Case "0"    'TCP
'''                        Disp1_sock(5).Protocol = sckTCPProtocol
'''                        Gate1_sock(5).Protocol = sckTCPProtocol
'''                    Case "1"    'UDP Only Send
'''                        Disp1_sock(5).Protocol = sckUDPProtocol
'''                        Disp1_sock(5).RemoteHost = LANE6_DispIP
'''                        Disp1_sock(5).RemotePort = LANE6_DispPort
'''                        Gate1_sock(5).Protocol = sckUDPProtocol
'''                        Gate1_sock(5).RemoteHost = LANE6_DeviceIP
'''                        Gate1_sock(5).RemotePort = LANE6_RelayPort
'''                End Select
        '    End If
    'End If
    
    
    'If (LANE1_YN = "Y" And LANE1_LPRMode = "2") Or (LANE2_YN = "Y" And LANE2_LPRMode = "2") Or (LANE3_YN = "Y" And LANE3_LPRMode = "2") Or (LANE4_YN = "Y" And LANE4_LPRMode = "2") Then
'        Call Hook
    'End If
    
    'Remote UDP ����
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
        FrmTcpServer.HomeSock.Close
        FrmTcpServer.HomeSock.Protocol = sckUDPProtocol
        FrmTcpServer.HomeSock.RemoteHost = HomeNet_IP
        FrmTcpServer.HomeSock.RemotePort = HomeNet_Port
        'Call RunHomeNet '�ּ�ó��:��Ŀ���� ���������� �̵��ϸ鼭 ȭ��ڷ� �������
    End If
    If (MVR_YN = "Y") Then
        MvrSock.Protocol = sckUDPProtocol
        MvrSock.RemoteHost = MVR_IP
        MvrSock.RemotePort = MVR_Port
    End If
    
    If (Glo_FreepassS_YN = "Y") Then
        FreepassS_sock.Protocol = sckUDPProtocol
        FreepassS_sock.RemoteHost = Glo_FreepassS_IP
        FreepassS_sock.RemotePort = Glo_FreepassS_Port
    End If
    If (Glo_FreepassR_YN = "Y") Then
        FreepassR_sock.Protocol = sckUDPProtocol
        FreepassR_sock.LocalPort = Glo_FreepassR_Port
        FreepassR_sock.Bind
    End If

    
    'TCP Server
'    MobileR_Sock.Protocol = sckTCPProtocol
'    MobileR_Sock.LocalPort = 30000
'    MobileR_Sock.Listen
    
    'UDP Server
'    MobileR_Sock.Protocol = sckUDPProtocol
'    MobileR_Sock.LocalPort = 30000
'    MobileR_Sock.Bind
    
    If (Glo_ParkFullLIGHT_YN = "Y") Then
        If (Glo_ParkFullLight_DispMode = "0") Then
            ParkFullLightS_sock.Protocol = sckTCPProtocol
            Glo_ParkFullLIGHT_PORT = 5000
        Else
            Glo_ParkFullLIGHT_PORT = 5108
            ParkFullLightS_sock.Protocol = sckUDPProtocol
            ParkFullLightS_sock.RemoteHost = Glo_ParkFullLIGHT_IP
            ParkFullLightS_sock.RemotePort = Glo_ParkFullLIGHT_PORT
            
        End If
        
    End If
    
    
    '���ܱ�ݱ� ������ſ� UDP���� ���� ����
    If (Glo_Lane1_GateClose_YN = "Y") Then
        Winsock_GateAgentR(0).Close
        Winsock_GateAgentR(0).Protocol = sckUDPProtocol
        Winsock_GateAgentR(0).LocalPort = 30201
        Winsock_GateAgentR(0).Bind
    End If
    If (Glo_Lane2_GateClose_YN = "Y") Then
        Winsock_GateAgentR(1).Close
        Winsock_GateAgentR(1).Protocol = sckUDPProtocol
        Winsock_GateAgentR(1).LocalPort = 30202
        Winsock_GateAgentR(1).Bind
    End If
    If (Glo_Lane3_GateClose_YN = "Y") Then
        Winsock_GateAgentR(2).Close
        Winsock_GateAgentR(2).Protocol = sckUDPProtocol
        Winsock_GateAgentR(2).LocalPort = 30203
        Winsock_GateAgentR(2).Bind
    End If
    If (Glo_Lane4_GateClose_YN = "Y") Then
        Winsock_GateAgentR(3).Close
        Winsock_GateAgentR(3).Protocol = sckUDPProtocol
        Winsock_GateAgentR(3).LocalPort = 30204
        Winsock_GateAgentR(3).Bind
    End If
    If (Glo_Lane5_GateClose_YN = "Y") Then
        Winsock_GateAgentR(4).Close
        Winsock_GateAgentR(4).Protocol = sckUDPProtocol
        Winsock_GateAgentR(4).LocalPort = 30205
        Winsock_GateAgentR(4).Bind
    End If
    If (Glo_Lane6_GateClose_YN = "Y") Then
        Winsock_GateAgentR(5).Close
        Winsock_GateAgentR(5).Protocol = sckUDPProtocol
        Winsock_GateAgentR(5).LocalPort = 30206
        Winsock_GateAgentR(5).Bind
    End If

    '���ܱ�ݱ� ������ſ� TCP���� ���� ����
'    If (Glo_Lane1_GateClose_YN = "Y") Then
'        Call Server_GateAgentR(0).StartServer(30201, "0.0.0.0")
'    End If
'    If (Glo_Lane2_GateClose_YN = "Y") Then
'        Call Server_GateAgentR(1).StartServer(30202, "0.0.0.0")
'    End If
'    If (Glo_Lane3_GateClose_YN = "Y") Then
'        Call Server_GateAgentR(2).StartServer(30203, "0.0.0.0")
'    End If
'    If (Glo_Lane4_GateClose_YN = "Y") Then
'        Call Server_GateAgentR(3).StartServer(30204, "0.0.0.0")
'    End If
'    If (Glo_Lane5_GateClose_YN = "Y") Then
'        Call Server_GateAgentR(4).StartServer(30205, "0.0.0.0")
'    End If
'    If (Glo_Lane6_GateClose_YN = "Y") Then
'        Call Server_GateAgentR(5).StartServer(30206, "0.0.0.0")
'    End If
   
   
    WinsockS_Devices.Protocol = sckUDPProtocol
    WinsockS_Devices.RemoteHost = "192.168.100.100"
    WinsockS_Devices.RemotePort = 45678
    
    
    
Exit Sub

Err_Proc:
    MsgBox ("[SetCommunication]  " & Err.Description)
    Call DataLogger(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    [SetCommunication]  " & Err.Description)

End Sub




Private Sub Form_Load()
    
    Call DataLogger("[HOST]    ���α׷� ����!!")
'''    If (Glo_ReANPR_YN = "Y") Then
'''        Call Rec_EngineOpen
'''    End If


    Call Certify_PC '��ϵ� PC���� üũ
    Glo_PartName = Glo_IPAddr
    
    
    Call Form_Activate
    
End Sub


'�����ü�� ����
Private Sub Certify()

    Dim rs As ADODB.Recordset
    Dim qry As String
    Dim LockDate As String
    Dim UnLockDate As String

On Error GoTo Err_p


    Call GetClienKey(Glo_PhyHDDKey)
    Glo_Certify = enumCertify.eCertNoTry
    
    Set rs = New ADODB.Recordset
    qry = "SELECT LockDate, UnLockDate FROM tb_Certify WHERE HASHCODE = '" & Glo_PhyHDDKey & "' "

    If (DataBaseQuery(rs, adoConn, qry, NWERR_GATE_STAY) = False) Then
        Call DebugLogger("[CERTIFY]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
        Exit Sub
    End If

''''    If rs.EOF Then
''''        Glo_Certify = True   '�ʱⰪ�� ������ü�� �������� �ʾ����Ƿ� �������� ó��
''''        cmd_Certify.Caption = "����"
''''        cmd_Certify.ToolTipText = "����Ű�� ������ �ִ� �ý��۰����ڸ� ����ϼ���"
''''        txt_CertifyKey.Visible = False
''''        lbl_CertifyLimitDate.Visible = False
''''
''''    Else
    

'''                LockDate = "" & rs!LockDate
'''                UnLockDate = "" & rs!UnLockDate
'''
'''                If (Len(UnLockDate) > 0) Then    '�����޾��� ���
'''                    Glo_Certify = True
'''
'''                    cmd_Certify.Caption = "�����Ϸ�"
'''                    cmd_Certify.ToolTipText = ""
'''                    cmd_Certify.Enabled = False
'''                    txt_CertifyKey.Visible = False
'''                    lbl_CertifyLimitDate.Visible = False
'''
'''                ElseIf (Len(LockDate) > 0) Then
'''                    Glo_Certify = False
'''                    Glo_Cert_LimitDate = DateAdd("m", Glo_Cert_Month, LockDate)
'''                    Glo_Cert_NoticeSDate = DateAdd("m", Glo_Cert_Month - 1, LockDate) '����Ⱓ �ȳ� ������(������ 1���� ��)
'''
'''                    cmd_Certify.Caption = "�����ʿ�"
'''                    cmd_Certify.ToolTipText = "�ݵ�� ������ �������� ������������. ���������ĺ��� ���ܱⰡ ���������� �ʽ��ϴ�"
'''                    txt_CertifyKey.Visible = True
'''                    lbl_CertifyLimitDate.Caption = "����Ⱓ:" & Glo_Cert_LimitDate
'''                    lbl_CertifyLimitDate.Visible = True
'''
'''                ElseIf (Len(LockDate) <= 0) Then
'''                    Glo_Certify = True
'''                End If
'''
'''        End If


        Do While Not (rs.EOF)

                LockDate = "" & rs!LockDate
                UnLockDate = "" & rs!UnLockDate
                
                
                If (Len(LockDate) > 0) Then    '������ư ������ ���
                    
                    If (Len(UnLockDate) > 0) Then '���� �Ϸ��� ���
                        Glo_Certify = enumCertify.eCertOK

                    Else
                        Glo_Certify = enumCertify.eCertTry

                    End If

                    Exit Do
                End If
                
                rs.MoveNext
        Loop
        
        
        If (Glo_Certify = enumCertify.eCertNoTry) Then
            cmd_Certify.Caption = "����"
            cmd_Certify.ToolTipText = "����Ű�� ������ �ִ� �ý��۰����ڸ� ����ϼ���"
            txt_CertifyKey.Visible = False
            lbl_CertifyLimitDate.Visible = False
        
        ElseIf (Glo_Certify = enumCertify.eCertTry) Then
            Glo_Cert_LimitDate = DateAdd("m", Glo_Cert_Month, LockDate)
            Glo_Cert_NoticeSDate = DateAdd("m", Glo_Cert_Month - 1, LockDate) '����Ⱓ �ȳ� ������(������ 1���� ��)
            
            cmd_Certify.Caption = "�����ʿ�"
            cmd_Certify.ToolTipText = "�ݵ�� ������ �������� ������������. ���������ĺ��� ���ܱⰡ ���������� �ʽ��ϴ�"
            txt_CertifyKey.Visible = True
            lbl_CertifyLimitDate.Caption = "����Ⱓ:" & Glo_Cert_LimitDate
            lbl_CertifyLimitDate.Visible = True
        
        ElseIf (Glo_Certify = enumCertify.eCertOK) Then
            cmd_Certify.Caption = "�����Ϸ�"
            cmd_Certify.ToolTipText = ""
            cmd_Certify.Enabled = False
            txt_CertifyKey.Visible = False
            lbl_CertifyLimitDate.Visible = False
        End If
    
    Set rs = Nothing
    
    
    Exit Sub
    
Err_p:
    Set rs = Nothing
    Call DebugLogger("[CERTIFY] Cert Res:" & Glo_Certify & ", Limit Date: " & Glo_Cert_LimitDate & ", Err: " & Err.Description)
End Sub


Private Sub FreepassR_sock_DataArrival(ByVal bytesTotal As Long)
    Dim sdata As String
    Dim Tmp_Path As String
    Dim i, gateNo As Integer
    Dim sPassKind As String
    Dim sPassDate As String
    Dim sStrLine() As String
    
On Error GoTo Err_p

    If (bytesTotal > 500) Then
        'DebugLogger ("RemoteR ������ �ʰ�����(������) : " & bytesTotal)
        Exit Sub
    End If
    
    
    FreepassR_sock.GetData sdata, , bytesTotal

    sdata = "" & sdata
    
    sStrLine() = Split(sdata, "_")
    
    gateNo = sStrLine(0)
    sPassKind = sStrLine(1)

    If (sPassKind = "FREEPASS" Or sPassKind = "TAXI" Or sPassKind = "NOWORK") Then ' �����н� ����
            Call DataLogger("FreepassR_sock" & "    " & sdata)
            Dim sYN As String
            sYN = sStrLine(2) ' �����н� Y or N
            
            '��ũ�� ���� ���� �б�
            If (Glo_Screen_No = 6) Then
                If (gateNo < Glo_Screen_No) Then
                    Call G6_23_Freepass(sPassKind, gateNo, sYN)
                End If
            ElseIf (Glo_Screen_No = 4) Then
                If (gateNo < Glo_Screen_No) Then
                    Call G4Mini_4IN_Freepass(sPassKind, gateNo, sYN)
                End If
            ElseIf (Glo_Screen_No = 2) Then
                If (gateNo < Glo_Screen_No) Then
                    Call Jung_Freepass(sPassKind, gateNo, sYN)
                End If
            ElseIf (Glo_Screen_No = 1) Then
                If (gateNo < Glo_Screen_No) Then
                    Call G1_Freepass(sPassKind, gateNo, sYN)
                End If
            End If
    End If

Exit Sub

Err_p:
    Call DataLogger(" [FreePassR_sock UDP DataArrival]  " & Err.Description)

End Sub


Private Sub Gate1_sock_SendProgress(Index As Integer, ByVal BytesSent As Long, ByVal BytesRemaining As Long)
    'Call DataLogger("[GATE TCP/IP ����] ������ LANE" & Index & ": bytesSent:" & BytesSent & ", Remain:" & BytesRemaining)
End Sub

Private Sub GateTimer_Timer(Index As Integer)

    If (GateTimer_First(Index)) Then
        GateTimer_First(Index) = False
        Exit Sub
    End If
    If (Gate_ACK(Index) = False) Then
        If (GlO_SendCnt(Index) >= 2) Then
            GateTimer(Index).Enabled = False
            GlO_SendCnt(Index) = 0
        Else
            
            If (Gate1_sock(Index).State <> sckClosed) Then
                Gate1_sock(Index).Close
            End If
            
            Select Case Index
                   Case 0
                        Gate1_sock(Index).Connect LANE1_DeviceIP, LANE1_RelayPort
                        If (GlO_GateRNum(Index) = 0) Then
                            Call DataLogger("[GATE TCP/IP ����]  �õ� IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort)
                        Else
                            Call DataLogger("[Get Frame TCP/IP ����]  �õ� IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort)
                        End If
                   Case 1
                        Gate1_sock(Index).Connect LANE2_DeviceIP, LANE2_RelayPort
                        If (GlO_GateRNum(Index) = 0) Then
                            Call DataLogger("[GATE TCP/IP ����]  �õ� IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort)
                        Else
                            Call DataLogger("[Get Frame TCP/IP ����]  �õ� IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort)
                        End If
                   Case 2
                        Gate1_sock(Index).Connect LANE3_DeviceIP, LANE3_RelayPort
                        If (GlO_GateRNum(Index) = 0) Then
                            Call DataLogger("[GATE TCP/IP ����]  �õ� IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_RelayPort)
                        Else
                            Call DataLogger("[Get Frame TCP/IP ����]  �õ� IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_RelayPort)
                        End If
                   Case 3
                        Gate1_sock(Index).Connect LANE4_DeviceIP, LANE4_RelayPort
                        If (GlO_GateRNum(Index) = 0) Then
                            Call DataLogger("[GATE TCP/IP ����]  �õ� IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort)
                        Else
                            Call DataLogger("[Get Frame TCP/IP ����]  �õ� IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort)
                        End If
                   Case 4
                        Gate1_sock(Index).Connect LANE5_DeviceIP, LANE5_RelayPort
                        If (GlO_GateRNum(Index) = 0) Then
                            Call DataLogger("[GATE TCP/IP ����]  �õ� IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_RelayPort)
                        Else
                            Call DataLogger("[Get Frame TCP/IP ����]  �õ� IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_RelayPort)
                        End If
                   Case 5
                        Gate1_sock(Index).Connect LANE6_DeviceIP, LANE6_RelayPort
                        If (GlO_GateRNum(Index) = 0) Then
                            Call DataLogger("[GATE TCP/IP ����]  �õ� IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_RelayPort)
                        Else
                            Call DataLogger("[Get Frame TCP/IP ����]  �õ� IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_RelayPort)
                        End If
            
            End Select
            
            GlO_SendCnt(Index) = GlO_SendCnt(Index) + 1
        End If
    Else
        GateTimer(Index).Enabled = False
    End If



End Sub


Private Sub MobileR_Sock_Close()
    MobileR_Sock.Close
    MobileR_Sock.Listen
    
    Call DebugLogger("MobileR_Sock Close : " & Format(Now, "yyyy-mm-dd hh:nn:ss") & Format(Timer * 1000 Mod 1000, " 000"))
End Sub

Private Sub MobileR_Sock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MobileR_Sock.Close
    MobileR_Sock.Listen
    
    Call DebugLogger("MobileR_Sock Error : " & Format(Now, "yyyy-mm-dd hh:nn:ss") & Format(Timer * 1000 Mod 1000, " 000"))
End Sub

Private Sub MobileR_Sock_ConnectionRequest(ByVal requestID As Long)
    MobileR_Sock.Close
    MobileR_Sock.Accept requestID
    
    Call DataLogger("Mobile Client : " & Format(Now, "yyyy-mm-dd hh:nn:ss") & Format(Timer * 1000 Mod 1000, " 000"))
End Sub

'MobileR_Sock TCP/IP ���
Private Sub MobileR_Sock_DataArrival(ByVal bytesTotal As Long)

    If (bytesTotal > 500) Then
        Exit Sub
    End If
    
    Dim sdata As String
    Dim sStrLine() As String
    
    MobileR_Sock.GetData sdata, , bytesTotal
    sdata = "" & sdata
    
    If (sdata = "") Then
        Exit Sub
    End If
    
    Call DataLogger("MobileR_Sock " & FormatNumber(bytesTotal, 0, , , vbTrue) & " bytes recieved." & "    " & sdata)

On Error GoTo Err_p

    
    MobileR_Sock.SendData "ACK1"
    
    
    sStrLine() = Split(sdata, "_")
    
    'Debug.Print sStrLine(0)
    'Debug.Print sStrLine(1)
    


    Exit Sub

Err_p:
    Call DataLogger(" [MobileR_Sock DataArrival]    " & Err.Description)
End Sub


Private Sub ParkFullLightS_sock_Connect()
    On Error GoTo Err_p
    Call DataLogger("[������ DISP TCP/IP ����] �Ϸ�")
    ParkFullLightS_sock.SendData GlO_ParkFullLight_BData
    Exit Sub
Err_p:
    Call DataLogger("[������ DISP TCP/IP ����] ���� : " & Err.Description)
    Call DebugLogger("[������ DISP TCP/IP ����] ���� : " & Err.Description)
End Sub
Private Sub ParkFullLightS_sock_SendComplete()
'    Call DataLogger("[������ TCP/IP ����] �Ϸ�")
End Sub
Private Sub ParkFullLightS_sock_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    Dim bData() As Byte
    Dim i As Integer

On Error GoTo Err_p

    ParkFullLightS_sock.GetData strData, , bytesTotal
'
'    If (Asc(strData) = 6) Then
'        Call DataLogger("[������ TCP/IP Rcv] " & "ACK")
'    Else
'        Call DataLogger("[������ TCP/IP Rcv] " & strData)
'    End If
'
    Call DataLogger("[������ DISP Rcv] " & FormatNumber(bytesTotal, 0, , , vbTrue) & " bytes recieved." & "    " & strData)
    Exit Sub
    
Err_p:
    Call DataLogger("[������ TCP/IP Rcv] ���� : " & Err.Description)
    Call DebugLogger("[������ TCP/IP Rcv] ���� : " & Err.Description)
End Sub

Private Sub ParkFullLightS_sock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error GoTo Err_p
    Call DataLogger("[������ TCP/IP ����] ���� : " & Description)
    Call DebugLogger("[������ TCP/IP ����] ����: " & Description)
    
    ParkFullLightS_sock.Close

    Call DataLogger("[������ TCP/IP ����] ��õ� IP = " & Glo_ParkFullLIGHT_IP & "    PORT = " & Glo_ParkFullLIGHT_PORT)
    Call DebugLogger("[������ TCP/IP ����] ��õ� IP = " & Glo_ParkFullLIGHT_IP & "    PORT = " & Glo_ParkFullLIGHT_PORT)
    ParkFullLightS_sock.Connect Glo_ParkFullLIGHT_IP, Glo_ParkFullLIGHT_PORT

    Exit Sub
Err_p:
    Call DebugLogger("[������ TCP/IP Err]  " & Err.Description)
End Sub


Private Sub Server_DataArrival(ByVal SckIndex As Integer, ByVal Data As String, ByVal bytesTotal As Long, ByVal RemoteIP As String, ByVal RemoteHost As String)

    Dim tmp_str As String
    
    If (bytesTotal > 500) Then
        'DebugLogger ("Server ������ �ʰ�����(������) : " & bytesTotal & ", Index:" & SckIndex)
        Exit Sub
    End If
    
    
    'LPR TCP/IP ���
    Call DataLogger("-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------")
        '������ ��ȣȭ
    tmp_str = DecodeNDE01(Data, "www.jawootek.com")
    Call DataLogger("Lane1 TCP Port " & FormatNumber(bytesTotal, 0, , , vbTrue) & " bytes recieved." & "    " & tmp_str)
    Server.SendData "ACK", SckIndex
    Call UDP_Proc(tmp_str)
    Exit Sub



Err_p:
    Call DataLogger(" [Server DataArrival]    " & Err.Description)
    

End Sub




Private Sub chk_RemoteYN_Click(Index As Integer)
    
'    Select Case Index
'
'        Case 0 ' ������ ����
'

            If chk_RemoteYN(Index).value = "1" Then
'                Glo_RemoteR_YN = "Y"
                Frame1(Index).Enabled = True
                If (Index = 1) Then
                    TxtSvrIp(Index).BackColor = &H80000005
                End If
            Else
'                Glo_RemoteR_YN = "N"
                Frame1(Index).Enabled = False
                If (Index = 1) Then
                    TxtSvrIp(Index).BackColor = &HE0E0E0
                End If
            End If
'            Call Put_Ini("System Config", "RemoteR_YN", Glo_RemoteR_YN)
'
'        Case 1 ' ������ �۽�
'
'            If chk_RemoteYN(Index).value = "1" Then
'                Glo_RemoteS_YN = "Y"
'                TxtSvrIp , RemoteS_IP
'            Else
'                Glo_RemoteS_YN = "N"
'                TxtSvrIp , RemoteS_IP
'
'            End If
'            Call Put_Ini("System Config", "RemoteS_YN", Glo_RemoteS_YN)
'
'    End Select

End Sub

Private Sub chk_UseYN_Click(Index As Integer)

    If (Index > Glo_Screen_No - 1) Then
        chk_UseYN(Index).value = "0"
    Else
        If chk_UseYN(Index).value = "1" Then
            Frame2(Index).Enabled = True
            cmb_Inout(Index).Enabled = True
            txt_GateName(Index).Enabled = True
            CmbScreen(Index).Enabled = True
        Else
            Frame2(Index).Enabled = False
            cmb_Inout(Index).Enabled = False
            txt_GateName(Index).Enabled = False
            CmbScreen(Index).Enabled = False
        End If
    End If
End Sub


Private Sub cmb_Disp1_Click(Index As Integer)
    Dim cmbIndex As Integer
    Dim cmbColor As Long
    Dim TxtColor As String


    If Cmb_Display.text = "������(Ǯ�÷�)_FW7" Then
    
        cmbIndex = cmb_Disp1(Index).ListIndex
        Select Case cmbIndex
            Case 0
                cmbColor = &HFF00& ' ���
            Case 1
                cmbColor = &HFF&   ' ����
            Case 2
                cmbColor = &H80C0FF ' Ȳ��
            Case 3
                cmbColor = &HFF0000 ' �Ķ�
            Case 4
                cmbColor = &H8B00FF ' ���ֻ�
            Case 5
                cmbColor = &HFFFF00 ' �ϴû�
            Case 6
                cmbColor = &HFFFFFF ' ���
        End Select
    
        txt_Disp1(Index).ForeColor = cmbColor
        
    ElseIf Cmb_Display.text = "������(Ǯ�÷�)" Then
    
        'cmbIndex = cmb_Disp1(Index).ListIndex
        Select Case cmb_Disp1(Index).text
            Case "��"
                cmbColor = &HFF00&   ' ���
            Case "��"
                cmbColor = &HFF     ' ����
            Case "Ȳ"
                cmbColor = &H80C0FF ' Ȳ��
        End Select
    
        txt_Disp1(Index).ForeColor = cmbColor
    Else
'''        cmbIndex = cmb_Disp1(Index).ListIndex
'''
'''        Select Case cmbIndex
'''            Case 0
'''                cmbColor = &HFF00&   ' ���
'''            Case 1
'''                cmbColor = &HFF  ' ����
'''            Case 2
'''                cmbColor = &H80C0FF ' Ȳ��
'''        End Select

        Select Case cmb_Disp1(Index).text
            Case "��"
                cmbColor = &HFF00&   ' ���
            Case "��"
                cmbColor = &HFF     ' ����
            Case "Ȳ"
                cmbColor = &H80C0FF ' Ȳ��
        End Select
    
        txt_Disp1(Index).ForeColor = cmbColor
    End If
End Sub

Private Sub cmb_Disp2_Click(Index As Integer)
    Dim cmbIndex As Integer
    Dim cmbColor As Long
    
    If Cmb_Display.text = "������(Ǯ�÷�)_FW7" Then

        'cmbIndex = cmb_Disp2(Index).ListIndex
        Select Case cmb_Disp2(Index).ListIndex
            Case 0
                cmbColor = &HFF00& ' ���
            Case 1
                cmbColor = &HFF&   ' ����
            Case 2
                cmbColor = &H80C0FF ' Ȳ��
            Case 3
                cmbColor = &HFF0000 ' �Ķ�
            Case 4
                cmbColor = &H8B00FF ' ���ֻ�
            Case 5
                cmbColor = &HFFFF00 ' �ϴû�
            Case 6
                cmbColor = &HFFFFFF ' ���
        End Select
        txt_Disp2(Index).ForeColor = cmbColor

    ElseIf Cmb_Display.text = "������(Ǯ�÷�)" Then

        'cmbIndex = cmb_Disp2(Index).ListIndex
        Select Case cmb_Disp2(Index).text
            Case "��"
                cmbColor = &HFF00&   ' ���
            Case "��"
                cmbColor = &HFF     ' ����
            Case "Ȳ"
                cmbColor = &H80C0FF ' Ȳ��
        End Select
        txt_Disp2(Index).ForeColor = cmbColor
        
    Else
'''        cmbIndex = cmb_Disp2(Index).ListIndex
'''
'''        Select Case cmbIndex
'''            Case 0
'''                cmbColor = &HFF00& ' ���
'''            Case 1
'''                cmbColor = &HFF&   ' ����
'''            Case 2
'''                cmbColor = &H80C0FF ' Ȳ��
'''        End Select
    
        Select Case cmb_Disp2(Index).text
            Case "��"
                cmbColor = &HFF00&   ' ���
            Case "��"
                cmbColor = &HFF     ' ����
            Case "Ȳ"
                cmbColor = &H80C0FF ' Ȳ��
        End Select
        
        txt_Disp2(Index).ForeColor = cmbColor
    End If
End Sub



'LPR ��Ź�� �����
Private Sub cmb_LPRMode_Click(Index As Integer)

    Select Case Index
        Case 0
            Select Case cmb_LPRMode(0).text
                Case "TCP"
                    txt_LPRPort(0).text = Trim(Server_Port)
                    txt_LPRPort(1).text = Trim(Server_Port)
                    txt_LPRPort(2).text = Trim(Server_Port)
                    txt_LPRPort(3).text = Trim(Server_Port)
                    txt_LPRPort(4).text = Trim(Server_Port)
                    txt_LPRPort(5).text = Trim(Server_Port)
                    
                Case Else
                    txt_LPRPort(0).text = Trim(LANE1_LPRPort)
                    txt_LPRPort(1).text = Trim(LANE2_LPRPort)
                    txt_LPRPort(2).text = Trim(LANE3_LPRPort)
                    txt_LPRPort(3).text = Trim(LANE4_LPRPort)
                    txt_LPRPort(4).text = Trim(LANE5_LPRPort)
                    txt_LPRPort(5).text = Trim(LANE6_LPRPort)
            End Select
        Case 1
            Select Case cmb_LPRMode(1).text
                Case "TCP"
                    txt_LPRPort(1).text = Trim(Server_Port)
                Case Else
                    txt_LPRPort(1).text = Trim(LANE2_LPRPort)
            End Select
        Case 2
            Select Case cmb_LPRMode(2).text
                Case "TCP"
                    txt_LPRPort(2).text = Trim(Server_Port)
                Case Else
                    txt_LPRPort(2).text = Trim(LANE3_LPRPort)
            End Select
        Case 3
            Select Case cmb_LPRMode(3).text
                Case "TCP"
                    txt_LPRPort(3).text = Trim(Server_Port)
                Case Else
                    txt_LPRPort(3).text = Trim(LANE4_LPRPort)
            End Select
        Case 4
            Select Case cmb_LPRMode(4).text
                Case "TCP"
                    txt_LPRPort(4).text = Trim(Server_Port)
                Case Else
                    txt_LPRPort(4).text = Trim(LANE5_LPRPort)
            End Select
        Case 5
            Select Case cmb_LPRMode(5).text
                Case "TCP"
                    txt_LPRPort(5).text = Trim(Server_Port)
                Case Else
                    txt_LPRPort(5).text = Trim(LANE6_LPRPort)
            End Select
    End Select

'����
'    If (cmb_LPRMode(Index).Text = "TCP") Then
'        txt_LPRPort(Index).Locked = True
'        txt_LPRPort(Index).BackColor = &HE0E0E0
'    Else
'        txt_LPRPort(Index).Locked = False
'        txt_LPRPort(Index).BackColor = &H80000005
'    End If

End Sub

Private Sub cmd_CapTest_Click(Index As Integer)
    
    If (Glo_LPRBoard = "�����" Or Glo_LPRBoard = "�ڵ��̳�") Then
        'Capture Test
        Call DataLogger("[Get Frame TEST]  Target Gate = " & Index)
        Call Relay_Out(1, Index)
        Call None_Delay_Time(1)

    Else
        Call DataLogger("Gate Up Test Error: " & Glo_LPRBoard)
    End If
    
End Sub

Private Sub cmd_EmgTest_Click(Index As Integer)
    
'    'Display Emg Test
'    If (Glo_LPRBoard = "�����") Then
'        Call DataLogger("[DISPLAY Emg TEST]  Target Gate = " & Index)
'        Call GL_Emergency("System Test", "System Test", 0, 30, 10, 1, 2, 1, Index)
'
'    ElseIf (Glo_LPRBoard = "�ڵ��̳�") Then
'        Call DataLogger("[DISPLAY Emg TEST]  Target Gate = " & Index)
'        'Call GL_Emergency("System Test", "System Test", 0, 30, 10, 1, 2, 1, Index) 'Old ������ ���� ����
'
'        If (Glo_Display_Direct = "����") Then
'            DoEvents
'            Call GL_Emergency_Horizontal("System Test", "System Test", 2, 1, Index) 'New ������ ���� ����
'        Else
'            DoEvents
'            'Call GL_Emergency_Vertical("�ϴö�����", "12345", 1, 1, Index) 'New ������ ���� ����
'            Call GL_Emergency_Vertical("Test", "System", 1, 1, Index) 'New ������ ���� ����
'        End If
'
'    Else
'        Call DataLogger("DISPLAY Emg TEST Error: " & Glo_LPRBoard)
'    End If
    
    
    If (Glo_Display = "������" Or Glo_Display = "������(Ǯ�÷�)") Then
        Call DataLogger("[DISPLAY Emg TEST]  Target Gate = " & Index)
        Call GL_Emergency("System Test", "System Test", 0, 30, 20, 1, 2, 1, Index)
        
    ElseIf (Glo_Display = "������(Ǯ�÷�)_FW7") Then
        Call DataLogger("[DISPLAY Emg TEST]  Target Gate = " & Index)
        
        If (Glo_Display_Direct = "����") Then
            DoEvents
            Call GL_Emergency_Horizontal("System Test", "System Test", enumDIS_COLOR2s.eGreen, enumDIS_COLOR2s.eYellow, Index)  'New ������ ���� ����
        Else
            DoEvents
            'Call GL_Emergency_Vertical("Test", "System", 1, 1, Index) 'New ������ ���� ����
            Call GL_Emergency_Vertical("Test", "System", enumDIS_COLOR2s.eGreen, enumDIS_COLOR2s.eGreen, Index) 'New ������ ���� ����
        End If
    End If

End Sub

Private Sub cmd_GateTest_Click(Index As Integer)
    
    'Gate Test
    If (Glo_LPRBoard = "�����" Or Glo_LPRBoard = "�ڵ��̳�") Then
        Call DataLogger("[GATE TEST]  Target Gate = " & Index)
        Call Relay_Out(0, Index)
    Else
        Call DataLogger("Gate Up Test Error: " & Glo_LPRBoard)
    End If

End Sub

Private Sub cmd_HomeTest_Click()
'    HomeNet_Dong = txt_Dong.Text
'    HomeNet_Ho = txt_Ho.Text
'    HomeNet_CarNo = "����01��1234"
'
'    HomeNet_Str = HomeNet_Dong & HomeNet_Ho & HomeNet_CarNo
'    FrmTcpServer.HomeSock.SendData (HomeNet_Str)
    'Call DataLogger("[HomeNet UDP ����]  DATA = " & HomeNet_Str)
End Sub

Private Sub cmd_NmlTest_Click(Index As Integer)

    Dim upColor As Byte
    Dim downColor As Byte
    
    If (Cmb_Display.text = "������(Ǯ�÷�)_FW7") Then
        Select Case cmb_Disp1(Index).text
            Case "��"
                upColor = enumDIS_COLOR2s.eRED
            Case "Ȳ"
                upColor = enumDIS_COLOR2s.eYellow
            Case "��"
                upColor = enumDIS_COLOR2s.eGreen
            Case "��"
                upColor = enumDIS_COLOR2s.eBLUE
            Case "��"
                upColor = enumDIS_COLOR2s.eWINE
            Case "��"
                upColor = enumDIS_COLOR2s.eSKY
            Case "��"
                upColor = enumDIS_COLOR2s.eWHITE
        End Select
        Select Case cmb_Disp2(Index).text
            Case "��"
                downColor = enumDIS_COLOR2s.eRED
            Case "Ȳ"
                downColor = enumDIS_COLOR2s.eYellow
            Case "��"
                downColor = enumDIS_COLOR2s.eGreen
            Case "��"
                downColor = enumDIS_COLOR2s.eBLUE
            Case "��"
                downColor = enumDIS_COLOR2s.eWINE
            Case "��"
                downColor = enumDIS_COLOR2s.eSKY
            Case "��"
                downColor = enumDIS_COLOR2s.eWHITE
        End Select
        
    ElseIf (Cmb_Display.text = "������(Ǯ�÷�)") Then 'Ȳ:2, ��:1, ��:0
        Select Case cmb_Disp1(Index).text
            Case "��"
                'upColor = enumDIS_COLORs.eRED
                upColor = 0
            Case "Ȳ"
                'upColor = enumDIS_COLORs.eYellow
                upColor = 2
            Case "��"
                'upColor = enumDIS_COLORs.eGreen
                upColor = 1
        End Select
        Select Case cmb_Disp2(Index).text
            Case "��"
                'downColor = enumDIS_COLORs.eRED
                downColor = 0
            Case "Ȳ"
                'downColor = enumDIS_COLORs.eYellow
                downColor = 2
            Case "��"
                'downColor = enumDIS_COLORs.eGreen
                downColor = 1
        End Select
            
    Else
        Select Case cmb_Disp1(Index).text
            Case "��"
                upColor = 0
            Case "��"
                upColor = 1
            Case "Ȳ"
                upColor = 2
        End Select
        Select Case cmb_Disp2(Index).text
            Case "��"
                downColor = 0
            Case "��"
                downColor = 1
            Case "Ȳ"
                downColor = 2
        End Select
    End If
    
    '�Ϲݹ��� ����
    If (cmd_NmlShift(Index).Caption = "����") Then
        Glo_LANE_DISP_NML_SHIFT(Index) = enumDISP_NML_SHIFT.eSTOP
            
        If (Glo_Display_Direct = "����") Then
            txt_Disp1(Index) = LeftH(txt_Disp1(Index), Glo_DISP_COL * 2) '���� �������� 6���̹Ƿ� 12����(6x2) ������
            txt_Disp2(Index) = LeftH(txt_Disp2(Index), Glo_DISP_COL * 2) '���� �������� 6���̹Ƿ� 12����(6x2) ������
        ElseIf (Glo_Display_Direct = "����") Then
            txt_Disp1(Index) = Left(txt_Disp1(Index), Glo_DISP_COL) '���� �������� 6���̹Ƿ� 6���� ������
            txt_Disp2(Index) = Left(txt_Disp2(Index), Glo_DISP_COL) '���� �������� 6���̹Ƿ� 6���� ������
        End If
    End If
    
    'Display Nomal Save
    'If (Glo_LPRBoard = "�����") Then
    If (Glo_Display = "������" Or Glo_Display = "������(Ǯ�÷�)") Then
        Call DataLogger("[DISPLAY Nomal]  Target Gate = " & Index)
        'Call GL_Nomal(txt_Disp1(Index), txt_Disp2(Index), 129, 70, 0, cmb_Disp1(Index).ListIndex, cmb_Disp2(Index).ListIndex, Index)
        Call GL_Nomal(txt_Disp1(Index), txt_Disp2(Index), 129, 70, 0, upColor, downColor, Index)
        
    'ElseIf (Glo_LPRBoard = "�ڵ��̳�") Then
    ElseIf (Glo_Display = "������(Ǯ�÷�)_FW7") Then
        Call DataLogger("[DISPLAY Nomal Save]  Target Gate = " & Index)

        If (Glo_Display_Direct = "����") Then
            Call GL_Nomal_Horizontal(txt_Disp1(Index), txt_Disp2(Index), 129, cmb_DispShiftSpeed(Index).text * 10, 0, upColor, downColor, Index, Glo_LANE_DISP_NML_SHIFT(Index)) '������ ����ǥ��(DABIT ������ ��� ��������:HEX), �������
        Else
            Call GL_Nomal_Vertical(txt_Disp1(Index), txt_Disp2(Index), 129, cmb_DispShiftSpeed(Index).text * 10, 0, upColor, downColor, Index, Glo_LANE_DISP_NML_SHIFT(Index)) '������ ����ǥ��(NEW ������ �߿���), �������
        End If

    Else
        Call DataLogger("DISPLAY Nomal TEST Error: " & Glo_LPRBoard)
        Exit Sub
    End If
    
    Call SaveNmlMsg(Index)
    

End Sub

Private Sub SaveNmlMsg(Index As Integer)
    Select Case Index
        Case 0
            LANE1_Disp1Msg = txt_Disp1(0)
            LANE1_Disp2Msg = txt_Disp2(0)
            LANE1_Disp1Color = CStr(cmb_Disp1(0).ListIndex)
            LANE1_Disp2Color = CStr(cmb_Disp2(0).ListIndex)
            Call Put_Ini("System Config", "LANE1_Disp1Msg", txt_Disp1(0))
            Call Put_Ini("System Config", "LANE1_Disp2Msg", txt_Disp2(0))
            Call Put_Ini("System Config", "LANE1_Disp1Color ", CStr(cmb_Disp1(0).ListIndex))
            Call Put_Ini("System Config", "LANE1_Disp2Color ", CStr(cmb_Disp2(0).ListIndex))
            Call Put_Ini("System Config", "LANE1_DispSpeed", CStr(cmb_DispShiftSpeed(0).ListIndex))
        
        Case 1
            LANE2_Disp1Msg = txt_Disp1(1)
            LANE2_Disp2Msg = txt_Disp2(1)
            LANE2_Disp1Color = CStr(cmb_Disp1(1).ListIndex)
            LANE2_Disp2Color = CStr(cmb_Disp2(1).ListIndex)
            Call Put_Ini("System Config", "LANE2_Disp1Msg", txt_Disp1(1))
            Call Put_Ini("System Config", "LANE2_Disp2Msg", txt_Disp2(1))
            Call Put_Ini("System Config", "LANE2_Disp1Color ", CStr(cmb_Disp1(1).ListIndex))
            Call Put_Ini("System Config", "LANE2_Disp2Color ", CStr(cmb_Disp2(1).ListIndex))
            Call Put_Ini("System Config", "LANE2_DispSpeed", CStr(cmb_DispShiftSpeed(1).ListIndex))
        
        Case 2
            LANE3_Disp1Msg = txt_Disp1(2)
            LANE3_Disp2Msg = txt_Disp2(2)
            LANE3_Disp1Color = CStr(cmb_Disp1(2).ListIndex)
            LANE3_Disp2Color = CStr(cmb_Disp2(2).ListIndex)
            Call Put_Ini("System Config", "LANE3_Disp1Msg", txt_Disp1(2))
            Call Put_Ini("System Config", "LANE3_Disp2Msg", txt_Disp2(2))
            Call Put_Ini("System Config", "LANE3_Disp1Color ", CStr(cmb_Disp1(2).ListIndex))
            Call Put_Ini("System Config", "LANE3_Disp2Color ", CStr(cmb_Disp2(2).ListIndex))
            Call Put_Ini("System Config", "LANE3_DispSpeed", CStr(cmb_DispShiftSpeed(2).ListIndex))
        
        Case 3
            LANE4_Disp1Msg = txt_Disp1(3)
            LANE4_Disp2Msg = txt_Disp2(3)
            LANE4_Disp1Color = CStr(cmb_Disp1(3).ListIndex)
            LANE4_Disp2Color = CStr(cmb_Disp2(3).ListIndex)
            Call Put_Ini("System Config", "LANE4_Disp1Msg", txt_Disp1(3))
            Call Put_Ini("System Config", "LANE4_Disp2Msg", txt_Disp2(3))
            Call Put_Ini("System Config", "LANE4_Disp1Color ", CStr(cmb_Disp1(3).ListIndex))
            Call Put_Ini("System Config", "LANE4_Disp2Color ", CStr(cmb_Disp2(3).ListIndex))
            Call Put_Ini("System Config", "LANE4_DispSpeed", CStr(cmb_DispShiftSpeed(3).ListIndex))
            
        Case 4
            LANE5_Disp1Msg = txt_Disp1(4)
            LANE5_Disp2Msg = txt_Disp2(4)
            LANE5_Disp1Color = CStr(cmb_Disp1(4).ListIndex)
            LANE5_Disp2Color = CStr(cmb_Disp2(4).ListIndex)
            Call Put_Ini("System Config", "LANE5_Disp1Msg", txt_Disp1(4))
            Call Put_Ini("System Config", "LANE5_Disp2Msg", txt_Disp2(4))
            Call Put_Ini("System Config", "LANE5_Disp1Color ", CStr(cmb_Disp1(4).ListIndex))
            Call Put_Ini("System Config", "LANE5_Disp2Color ", CStr(cmb_Disp2(4).ListIndex))
            Call Put_Ini("System Config", "LANE5_DispSpeed", CStr(cmb_DispShiftSpeed(4).ListIndex))
    
        Case 5
            LANE6_Disp1Msg = txt_Disp1(5)
            LANE6_Disp2Msg = txt_Disp2(5)
            LANE6_Disp1Color = CStr(cmb_Disp1(5).ListIndex)
            LANE6_Disp2Color = CStr(cmb_Disp2(5).ListIndex)
            Call Put_Ini("System Config", "LANE6_Disp1Msg", txt_Disp1(5))
            Call Put_Ini("System Config", "LANE6_Disp2Msg", txt_Disp2(5))
            Call Put_Ini("System Config", "LANE6_Disp1Color ", CStr(cmb_Disp1(5).ListIndex))
            Call Put_Ini("System Config", "LANE6_Disp2Color ", CStr(cmb_Disp2(5).ListIndex))
            Call Put_Ini("System Config", "LANE6_DispSpeed", CStr(cmb_DispShiftSpeed(5).ListIndex))
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
            LANE1_Inout = cmb_Inout(0).text
            LANE1_Name = Trim(txt_GateName(0))
            LANE1_LPRMode = cmb_LPRMode(0).ListIndex
            LANE1_LPRIP = Trim(txt_LPRIP(0))
            LANE1_LPRPort = Trim(txt_LPRPort(0))
            LANE1_DeviceMode = cmb_DeviceMode(0).ListIndex
            LANE1_DeviceIP = Trim(txt_DeviceIP(0))
            LANE1_DispIP = Trim(txt_DispIP(0))
            LANE1_DisplayMode = cmb_DisplayMode(0).ListIndex
            LANE1_DispPort = Trim(txt_DispPort(0))
            LANE1_RelayPort = Trim(txt_RelayPort(0))
            
            Call Put_Ini("System Config", "LPRMode", LANE1_LPRMode)
            Call Put_Ini("System Config", "DisplayMode ", LANE1_DisplayMode)
            Call Put_Ini("System Config", "DeviceMode ", LANE1_DeviceMode)
            
            'LANE1_DispComPort = cmb_DispComPort(0).Text
            'LANE1_RelayComPort = cmb_RelayComPort(0).Text
            Call Put_Ini("System Config", "LANE1_YN ", LANE1_YN)
            Call Put_Ini("System Config", "LANE1_INOUT ", LANE1_Inout)
            Call Put_Ini("System Config", "LANE1_Name ", LANE1_Name)
            'Call Put_Ini("System Config", "LANE1_LPRMode ", LANE1_LPRMode)
            Call Put_Ini("System Config", "LANE1_LPRIP ", LANE1_LPRIP)
            'Call Put_Ini("System Config", "LANE1_LPRPort ", CStr(LANE1_LPRPort))
            'Call Put_Ini("System Config", "LANE1_DeviceMode ", LANE1_DeviceMode)
            Call Put_Ini("System Config", "LANE1_DeviceIP ", LANE1_DeviceIP)
            Call Put_Ini("System Config", "LANE1_DispIP ", LANE1_DispIP)
            'Call Put_Ini("System Config", "LANE1_DispPort ", CStr(LANE1_DispPort))
            'Call Put_Ini("System Config", "LANE1_RelayPort ", CStr(LANE1_RelayPort))
            'Call Put_Ini("System Config", "LANE1_DispComPort ", CStr(LANE1_DispComPort))
            'Call Put_Ini("System Config", "LANE1_RelayComPort ", CStr(LANE1_RelayComPort))
    
        Case 1
            If chk_UseYN(1).value = "1" Then
                LANE2_YN = "Y"
            Else
                LANE2_YN = "N"
            End If
            LANE2_Inout = cmb_Inout(1).text
            LANE2_Name = Trim(txt_GateName(1))
            LANE2_LPRMode = cmb_LPRMode(1).ListIndex
            LANE2_LPRIP = Trim(txt_LPRIP(1))
            LANE2_LPRPort = Trim(txt_LPRPort(1))
            LANE2_DeviceMode = cmb_DeviceMode(1).ListIndex
            LANE2_DeviceIP = Trim(txt_DeviceIP(1))
            LANE2_DispIP = Trim(txt_DispIP(1))
            LANE2_DisplayMode = cmb_DisplayMode(1).ListIndex
            LANE2_DispIP = Trim(txt_DispIP(1))
            LANE2_DispPort = Trim(txt_DispPort(1))
            LANE2_RelayPort = Trim(txt_RelayPort(1))
            Call Put_Ini("System Config", "LANE2_YN ", LANE2_YN)
            Call Put_Ini("System Config", "LANE2_INOUT ", LANE2_Inout)
            Call Put_Ini("System Config", "LANE2_Name ", LANE2_Name)
            Call Put_Ini("System Config", "LANE2_LPRIP ", LANE2_LPRIP)
            Call Put_Ini("System Config", "LANE2_DeviceIP ", LANE2_DeviceIP)
            Call Put_Ini("System Config", "LANE2_DispIP ", LANE2_DispIP)
    
        Case 2
            If chk_UseYN(2).value = "1" Then
                LANE3_YN = "Y"
            Else
                LANE3_YN = "N"
            End If
            LANE3_Inout = cmb_Inout(2).text
            LANE3_Name = Trim(txt_GateName(2))
            LANE3_LPRMode = cmb_LPRMode(2).ListIndex
            LANE3_LPRIP = Trim(txt_LPRIP(2))
            LANE3_LPRPort = Trim(txt_LPRPort(2))
            LANE3_DeviceMode = cmb_DeviceMode(2).ListIndex
            LANE3_DeviceIP = Trim(txt_DeviceIP(2))
            LANE3_DispIP = Trim(txt_DispIP(2))
            LANE3_DisplayMode = cmb_DisplayMode(2).ListIndex
            LANE3_DispIP = Trim(txt_DispIP(2))
            LANE3_DispPort = Trim(txt_DispPort(2))
            LANE3_RelayPort = Trim(txt_RelayPort(2))
            Call Put_Ini("System Config", "LANE3_YN ", LANE3_YN)
            Call Put_Ini("System Config", "LANE3_INOUT ", LANE3_Inout)
            Call Put_Ini("System Config", "LANE3_Name ", LANE3_Name)
            Call Put_Ini("System Config", "LANE3_LPRIP ", LANE3_LPRIP)
            Call Put_Ini("System Config", "LANE3_DeviceIP ", LANE3_DeviceIP)
            Call Put_Ini("System Config", "LANE3_DispIP ", LANE3_DispIP)
    
        Case 3
            If chk_UseYN(3).value = "1" Then
                LANE4_YN = "Y"
            Else
                LANE4_YN = "N"
            End If
            LANE4_Inout = cmb_Inout(3).text
            LANE4_Name = Trim(txt_GateName(3))
            LANE4_LPRMode = cmb_LPRMode(3).ListIndex
            LANE4_LPRIP = Trim(txt_LPRIP(3))
            LANE4_LPRPort = Trim(txt_LPRPort(3))
            LANE4_DeviceMode = cmb_DeviceMode(3).ListIndex
            LANE4_DeviceIP = Trim(txt_DeviceIP(3))
            LANE4_DisplayMode = cmb_DisplayMode(3).ListIndex
            LANE4_DispIP = Trim(txt_DispIP(3))
            LANE4_DispPort = Trim(txt_DispPort(3))
            LANE4_RelayPort = Trim(txt_RelayPort(3))
            Call Put_Ini("System Config", "LANE4_YN ", LANE4_YN)
            Call Put_Ini("System Config", "LANE4_INOUT ", LANE4_Inout)
            Call Put_Ini("System Config", "LANE4_Name ", LANE4_Name)
            Call Put_Ini("System Config", "LANE4_LPRIP ", LANE4_LPRIP)
            Call Put_Ini("System Config", "LANE4_DeviceIP ", LANE4_DeviceIP)
            Call Put_Ini("System Config", "LANE4_DispIP ", LANE4_DispIP)
            
        Case 4
            If chk_UseYN(4).value = "1" Then
                LANE5_YN = "Y"
            Else
                LANE5_YN = "N"
            End If
            LANE5_Inout = cmb_Inout(4).text
            LANE5_Name = Trim(txt_GateName(4))
            LANE5_LPRMode = cmb_LPRMode(4).ListIndex
            LANE5_LPRIP = Trim(txt_LPRIP(4))
            LANE5_LPRPort = Trim(txt_LPRPort(4))
            LANE5_DeviceMode = cmb_DeviceMode(4).ListIndex
            LANE5_DeviceIP = Trim(txt_DeviceIP(4))
            LANE5_DisplayMode = cmb_DisplayMode(4).ListIndex
            LANE5_DispIP = Trim(txt_DispIP(4))
            LANE5_DispPort = Trim(txt_DispPort(4))
            LANE5_RelayPort = Trim(txt_RelayPort(4))
            Call Put_Ini("System Config", "LANE5_YN ", LANE5_YN)
            Call Put_Ini("System Config", "LANE5_INOUT ", LANE5_Inout)
            Call Put_Ini("System Config", "LANE5_Name ", LANE5_Name)
            Call Put_Ini("System Config", "LANE5_LPRIP ", LANE5_LPRIP)
            Call Put_Ini("System Config", "LANE5_DeviceIP ", LANE5_DeviceIP)
            Call Put_Ini("System Config", "LANE5_DispIP ", LANE5_DispIP)
            
        Case 5
            If chk_UseYN(5).value = "1" Then
                LANE6_YN = "Y"
            Else
                LANE6_YN = "N"
            End If
            LANE6_Inout = cmb_Inout(5).text
            LANE6_Name = Trim(txt_GateName(5))
            LANE6_LPRMode = cmb_LPRMode(5).ListIndex
            LANE6_LPRIP = Trim(txt_LPRIP(5))
            LANE6_LPRPort = Trim(txt_LPRPort(5))
            LANE6_DeviceMode = cmb_DeviceMode(5).ListIndex
            LANE6_DeviceIP = Trim(txt_DeviceIP(5))
            LANE6_DisplayMode = cmb_DisplayMode(5).ListIndex
            LANE6_DispIP = Trim(txt_DispIP(5))
            LANE6_DispPort = Trim(txt_DispPort(5))
            LANE6_RelayPort = Trim(txt_RelayPort(5))
            Call Put_Ini("System Config", "LANE6_YN ", LANE6_YN)
            Call Put_Ini("System Config", "LANE6_INOUT ", LANE6_Inout)
            Call Put_Ini("System Config", "LANE6_Name ", LANE6_Name)
            Call Put_Ini("System Config", "LANE6_LPRIP ", LANE6_LPRIP)
            Call Put_Ini("System Config", "LANE6_DeviceIP ", LANE6_DeviceIP)
            Call Put_Ini("System Config", "LANE6_DispIP ", LANE6_DispIP)
    End Select

End Sub

'Localhost Server Config Save
Private Sub cmd_Svr_Click()
    Dim i As Integer
        
    'Sever Refresh
    If (LANE1_LPRMode = "0") Then
        Call Server.StopServer
    End If
    Server_Port = Trim(txtPort)
    For i = 0 To 4
        If cmb_LPRMode(i).text = "TCP" Then
            txt_LPRPort(i).text = Trim(Server_Port)
            txt_LPRPort(i).Locked = True
            txt_LPRPort(i).BackColor = &HE0E0E0
        Else
            txt_LPRPort(i).Locked = False
            txt_LPRPort(i).BackColor = &H80000005
        End If
    Next i
    If (LANE1_LPRMode = "0") Then
        Call Server.StartServer(Server_Port, Server.ServerIP)
    End If

    
End Sub




'����
Private Sub Command3_Click()
    Dim PROC As Integer
    
On Error GoTo Err_p
    
    PROC = Shell("explorer.exe /n,/e", vbNormalFocus)
    Exit Sub
Err_p:
    Call DataLogger("Ž�����������:" & Err.Description)
End Sub
'����
Private Sub Command4_Click()
'    Dim PROC As Integer
'    PROC = Shell("C:\Program Files\WIZnet\WIZ12xSR Configuration Tool\WIZ12xSR_CFG.exe", vbNormalFocus)
 '   frmSEGConf.Show 0
End Sub
'����
Private Sub Command5_Click()
    Dim PROC As Integer
On Error GoTo Err_p

    PROC = Shell("notepad .\winpark.ini", vbNormalFocus)
    Exit Sub
Err_p:
    Call DataLogger("����â�������:" & Err.Description)
End Sub

Private Sub Command6_Click()
On Error GoTo Err_p
    'FrmExtend.Show 1
    FrmExtend.Show 0
    Exit Sub
Err_p:
    Call DataLogger("���μ���ȭ�����:" & Err.Description)
End Sub

Private Sub Command7_Click()
    Dim PROC As Integer
On Error GoTo Err_p
    
    PROC = Shell("cmd.exe", vbNormalFocus)
    Exit Sub
Err_p:
    Call DataLogger("���â�������:" & Err.Description)
End Sub

Public Sub CloseAllSock()

    Dim i As Integer
    
    Call Server.StopServer
    Call Server_WebDC.StopServer
    
    For i = 0 To MAX_LANE_COUNT - 1
        LPR1_sock(i).Close
        Disp1_sock(i).Close
        Gate1_sock(i).Close
        LPR_Send_sock(i).Close
'        Set Gate1_UniSock(i) = Nothing
        Reset_sock(i).Close

    Next i


    RemoteS_sock.Close
    RemoteR_sock.Close
    FreepassS_sock.Close
    FreepassR_sock.Close
    HomeSock.Close
    MvrSock.Close
    DeviceR_sock.Close
    ParkFullLightS_sock.Close

    MobileR_Sock.Close
    
    Call Server_GateAgentR(0).StopServer
    Call Server_GateAgentR(1).StopServer
    Call Server_GateAgentR(2).StopServer
    Call Server_GateAgentR(3).StopServer
    Call Server_GateAgentR(4).StopServer
    Call Server_GateAgentR(5).StopServer
    
    Winsock_GateAgentR(0).Close
    Winsock_GateAgentR(1).Close
    Winsock_GateAgentR(2).Close
    Winsock_GateAgentR(3).Close
    Winsock_GateAgentR(4).Close
    Winsock_GateAgentR(5).Close
    
    WinsockS_Devices.Close
'
End Sub

Private Sub Command8_Click()
    ListData.Clear
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call CloseAllSock
    
    If (Glo_ReANPR_YN = "Y") Then
        Call Rec_EngineClose
    End If
    
    Call DeleteCriticalSection(Glo_CS)
End Sub


Private Sub AppPW_Check()
    
    If (Glo_APP_CHG_DAY = 0) Then
        Exit Sub
    End If
    
    Dim rs As Recordset
    Dim qry As String
    Dim AppRegDate As Long

On Error GoTo Err_p

    Set rs = New ADODB.Recordset
    rs.Open "SELECT CAR_NO, APP_YN, APP_YES_DATE, APP_CERTIFY_DATE FROM tb_reg", adoConn
    Do While Not (rs.EOF)

        '����Ͼ� ��� + ����Ͼ� �α��� ��й�ȣ ����� �� ���
        If (rs!APP_YN = "Y" And Len("" & rs!APP_CERTIFY_DATE) = 0) Then

            AppRegDate = DateDiff("d", Left(rs!APP_YES_DATE, 10), Format(Now, "yyyy-mm-dd"))
            If (AppRegDate > Glo_APP_CHG_DAY) Then
                adoConn.Execute "UPDATE tb_reg     SET APP_YN='N', APP_PW='', APP_YES_DATE=Null, APP_CERTIFY_DATE=Null WHERE CAR_NO = '" & rs!CAR_NO & "'"
                adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('" & rs!CAR_NO & "', 'HOST','�� ��й�ȣ �������� �ʾƼ� ����� ������',''," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
                Call DataLogger("[AppPW Check]    " & rs!CAR_NO & "   �� ��й�ȣ �������� �ʾƼ� ����� ������")
            End If
        End If
        rs.MoveNext
    Loop
    Set rs = Nothing
    Exit Sub
Err_p:
    Set rs = Nothing
    Call DataLogger("[AppPW Check]    " & Err.Description)
    adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('', 'HOST','[AppPW Check] " & Err.Description & "',''," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
End Sub


Private Sub Server_GateAgentR_DataArrival(Index As Integer, ByVal SckIndex As Integer, ByVal Data As String, ByVal bytesTotal As Long, ByVal RemoteIP As String, ByVal RemoteHost As String)
    Dim sStrLine() As String
    Dim sLog As String
    
    On Error GoTo Err_p
    
    sLog = "LANE" & Index + 1 & " : " & Data
    adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('[GATE_AGENT]', 'HOST', '" & sLog & "',''," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
    Call DataLogger("[GATE ] " & sLog)
    Exit Sub
Err_p:
    Call DebugLogger("[Server_GateAgentR] DataArrival Error : " & Err.Description)
End Sub

Private Sub Server_WebDC_DataArrival(ByVal SckIndex As Integer, ByVal Data As String, ByVal bytesTotal As Long, ByVal RemoteIP As String, ByVal RemoteHost As String)
    Dim sStrLine() As String
    Dim sLog As String
    
    On Error GoTo Err_p
    
    ' 0_GATEOPEN_ī�丮��
    sStrLine() = Split(Data, "_")
    
    sLog = "���ܱ� ����: Lane" & Val(sStrLine(0)) + 1 & " : " & sStrLine(2)
    adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('[������]', 'WEBDC', '" & sLog & "',''," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
    Call DataLogger("[������] " & sLog)
    'Call DataLogger("[������] TCP Port " & FormatNumber(bytesTotal, 0, , , vbTrue) & " bytes recieved" & "    " & Data)
    
    Call cmd_GateTest_Click(Val(sStrLine(0))) '���ܱ� ����
    
    Server_WebDC.SendData Data & "_" & "ACK", SckIndex
    
    Exit Sub
Err_p:
    Call DebugLogger("[Server_WebDC] DataArrival Error : " & Err.Description)
End Sub

Private Sub Timer_Certify_Timer()

    On Error GoTo Err_p
    
    Timer_Certify.Enabled = False
    If (Glo_Certify = enumCertify.eCertTry) Then
        If (Format(Now, "nnss") = "0005") Then                          '1�ð�����
            If (Format(Now, "yyyy-mm-dd") > Glo_Cert_LimitDate) Then    '���������� �������, ���âǥ�� �� ���α׷� ��������
                
                Call DataLogger("[����] " & "�����Ⱓ ������ �������ϴ�. �������� �� �����ϼ���(������:" & Glo_Cert_LimitDate & ")")
                Msg_Box.Caption = "Parking System"
                Msg_Box.Label1.Caption = "������������ �������ϴ�. " & Chr$(13) & Chr$(10) & "�������� �� �����ϼ���" & Chr$(13) & Chr$(10) & "(������:" & Glo_Cert_LimitDate & ")"
                Msg_Box.Label2.Caption = "���������!!"
                Msg_Box.Show 1
                End '��������
                
            ElseIf (Format(Now, "yyyy-mm-dd") > Glo_Cert_NoticeSDate) Then '���������� 1���� �̳�
                Call DataLogger("[����] " & "������ �����Դϴ�. ���������� ���Ŀ��� ���ܱⰡ ���������� �ʽ��ϴ�(������:" & Glo_Cert_LimitDate & ")")
                Msg_Box.Caption = "Parking System"
                Msg_Box.Label1.Caption = "������ �����Դϴ�." & Chr$(13) & Chr$(10) & "������ ���Ŀ���, " & Chr$(13) & Chr$(10) & "���ܱⰡ ���������� �ʽ��ϴ�" & Chr$(13) & Chr$(10) & "(������:" & Glo_Cert_LimitDate & ")"
                Msg_Box.Label2.Caption = "���������!!"
                Msg_Box.Show 0
            End If
        End If
    End If
    Timer_Certify.Enabled = True
    Exit Sub

Err_p:
    If (Timer_Certify.Enabled = False) Then
        Timer_Certify.Enabled = True
    End If
End Sub

Private Sub Timer_Emerg_Vertical_Timer(Index As Integer)
    
'    FrmTcpServer.Timer_Emerg_Vertical(Index).Enabled = False
'    DoEvents
'    FrmTcpServer.Timer_Emerg_Vertical(Index).Interval = Glo_Emerg_Vertical_ToggleTime * 1000 '���� ms
'    FrmTcpServer.Timer_Emerg_Vertical(Index).Enabled = True
    
    '������ȣ ���
    If (Glo_Emerg_Vertical(Index).ToggleSelect = EnumEmergToggleOrder.enumCarNo) Then
            Glo_Emerg_Vertical(Index).ToggleSelect = EnumEmergToggleOrder.enumCarStat '���� ��ۿ��� ������ȣ ���
            Glo_Emerg_Vertical(Index).CarNoCount = Glo_Emerg_Vertical(Index).CarNoCount - 1
            Call GL_Emergency_Vertical(Glo_Emerg_Vertical(Index).CarNo1, Glo_Emerg_Vertical(Index).CarNo2, Glo_Emerg_Vertical(Index).CarNoColor1, Glo_Emerg_Vertical(Index).CarNoColor2, Index)
            
    'ó����� ���
    ElseIf (Glo_Emerg_Vertical(Index).ToggleSelect = EnumEmergToggleOrder.enumCarStat) Then
            Glo_Emerg_Vertical(Index).ToggleSelect = EnumEmergToggleOrder.enumCarNo   '���� ��ۿ��� ���� ���
            Glo_Emerg_Vertical(Index).CarStatCount = Glo_Emerg_Vertical(Index).CarStatCount - 1
            Call GL_Emergency_Vertical(Glo_Emerg_Vertical(Index).CarStat1, Glo_Emerg_Vertical(Index).CarStat2, Glo_Emerg_Vertical(Index).CarStatColor1, Glo_Emerg_Vertical(Index).CarStatColor2, Index)
            
    End If
    
    
    If (Glo_Emerg_Vertical(Index).CarNoCount <= 0 And Glo_Emerg_Vertical(Index).CarStatCount <= 0) Then
        Timer_Emerg_Vertical(Index).Enabled = False
    End If
    
    

End Sub

Private Sub Timer_ParkFullLight_Timer()
    '������
    If (Glo_ParkFullLIGHT_YN = "Y") Then
    
        Dim iValue As Long
        iValue = Int((Glo_ParkNow_Count / Glo_ParkFull_Count) * 100) ' %
        
        Glo_ParkFullLigth_Toggle = Glo_ParkFullLigth_Toggle Xor True
        
        
        '1 = "����"
        '2 = "�ʷ�"
        '3 = "���"
        '4 = "�Ķ�"
        '5 = "��ȫ"
        '6 = "�ϴ�"
        '7 = "���"
        If (Glo_ParkFullLigth_Toggle = True) Then

            '����
            If (iValue >= 100) Then
                Call GL_Nomal_ParkFullLight(Glo_ParkFullLIGHT_FULL, 1) '����ȭ��
            'ȥ��
            ElseIf (iValue >= Glo_ParkFullLIGHT_GUIDE) Then
                Call GL_Nomal_ParkFullLight(Glo_ParkFullLIGHT_BUSY, 3) '����ȭ��
            '����
            Else
                Call GL_Nomal_ParkFullLight(Glo_ParkFullLIGHT_EMPTY, 2)  '����ȭ��(�ʷ�)
            End If

        Else
            '����
            If (iValue >= 100) Then
                Call GL_Nomal_ParkFullLight(Glo_ParkFullLIGHT_FULL, 1) '����ȭ��
            'ȥ��
            ElseIf (iValue >= Glo_ParkFullLIGHT_GUIDE) Then
                Call GL_Nomal_ParkFullLight(CStr(Glo_ParkFull_Count - Glo_ParkNow_Count), 3) '����ȭ��
            '����
            Else
                Call GL_Nomal_ParkFullLight(CStr(Glo_ParkFull_Count - Glo_ParkNow_Count), 2) '����ȭ��(�ʷ�)
            End If
        End If

    End If
    
    
End Sub

'�湮���� ���� ����ó��
Private Sub Delete_GuestRegCar()
    Dim rs As ADODB.Recordset
    Dim sCarNo As String
    Dim sCarModel As String
    Dim sGuestRegEndDate As String
    Dim sNowDateTime As String
    Dim sName As String
    Dim sTel As String
    Dim sDond As String
    Dim sHo As String
    Dim sStartDate As String
    Dim sEndDate As String

    If (Glo_GuestReg_YN = "N") Then
        Exit Sub
    End If
    
    sNowDateTime = Format(Now, "yyyy-mm-dd hh:nn:ss")

    Set rs = New ADODB.Recordset
    rs.Open "SELECT car_no, CAR_MODEL, Driver_Name,Driver_Phone,Driver_Dept,Driver_Class,start_date, end_date FROM tb_guestReg where car_gubun='�湮����'", adoConn
    Do While Not (rs.EOF)
        sGuestRegEndDate = "" & Format(rs!END_DATE, "yyyy-mm-dd hh:nn:ss")
        'sGuestRegEndDate = Mid(rs!END_DATE, 1, 4) & Mid(rs!END_DATE, 6, 2) & Mid(rs!END_DATE, 8, 2)
        If (sGuestRegEndDate < sNowDateTime) Then

            sCarNo = "" & rs!CAR_NO
            sCarModel = "" & rs!CAR_MODEL
            sName = "" & rs!DRIVER_NAME
            sTel = "" & rs!DRIVER_PHONE
            sDond = "" & rs!DRIVER_DEPT
            sHo = "" & rs!DRIVER_CLASS
            sStartDate = Format(rs!START_DATE, "yyyy-mm-dd hh:nn:ss")
            sEndDate = Format(rs!END_DATE, "yyyy-mm-dd hh:nn:ss")

            adoConn.Execute "Delete from tb_guestReg WHERE CAR_NO = '" & sCarNo & "' AND END_DATE < '" & sNowDateTime & "' " '���� ������ �湮���� ���ڵ� ����
            adoConn.Execute "INSERT INTO tb_reg_log (CAR_NO,CAR_MODEL, CAR_GUBUN,DRIVER_NAME,DRIVER_PHONE,DRIVER_DEPT,DRIVER_CLASS,START_DATE,END_DATE,REG_DATE,ACTION_LOG,ACTION_ID) value ('" & sCarNo & "','" & sCarModel & "', '�湮����','" & sName & "','" & sTel & "','" & sDond & "','" & sHo & "','" & sStartDate & "','" & sEndDate & "','" & sNowDateTime & "','����','System') "

        End If

        rs.MoveNext
    Loop
    Set rs = Nothing
End Sub

'�����湮���� �ڵ� ���� ����(�ſ� 1ȸ ����)
Public Sub Init_GuestRegCar()
    
On Error GoTo Err_p
    
    If (Glo_GuestReg_YN = "N") Then
        Exit Sub
    End If
    
    Dim sQry As String
    Dim sLastUpdateDate As String
    Dim sNowDate As String
    Dim sLog As String
    
    sNowDate = Format(Now, "yyyy-mm-dd hh:nn:ss")
    
    sQry = "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('', 'HOST','[�����湮����] �ʱ�ȭ �����췯 ����','SYSTEM'," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
    adoConn.Execute sQry
    
    Set rs = New ADODB.Recordset
    rs.Open "SELECT ID, REG_DATE FROM tb_guestreg_admin WHERE USE_YN = 'Y'", adoConn
    Do While Not (rs.EOF)
        sLastUpdateDate = Left(rs!REG_DATE, 7)  'yyyy-mm
        
        If (sLastUpdateDate < Left(sNowDate, 7)) Then
            adoConn.Execute "UPDATE tb_guestreg_admin     SET NOWPARKCOUNT = 0, NOWPARKTIME = 0, REG_DATE = '" & sNowDate & "' WHERE USE_YN = 'Y'"
            
            sLog = "[�����湮����]    " & " �ʱ�ȭ ����(�湮Ƚ��=0, �����ð�=0) " & rs!ID
            Call DataLogger(sLog)
            sQry = "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('', 'HOST','" & sLog & "','SYSTEM'," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
            adoConn.Execute sQry
            
        End If
        
        rs.MoveNext
    Loop
    
    Set rs = Nothing
    
    sQry = "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('', 'HOST','[�����湮����] �ʱ�ȭ �����췯 ����','SYSTEM'," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
    adoConn.Execute sQry
    Exit Sub
    
Err_p:
    Set rs = Nothing
    Call DataLogger("[�����湮����]    " & "�ʱ�ȭ ����(E00009)" & " " & Err.Description)
    
End Sub

Public Sub Webdc_Charge_FreeAutoPoint()
    Dim rs As Recordset
    Dim rs2 As Recordset
    Dim sQry As String
    Dim sQry2 As String
    Dim bQryResult As Boolean
    Dim nAutoFreePoint As Integer
    Dim sSEQ, sID, sPSEQ, sPName As String
    Dim sLog As String
    Dim sStrLine() As String
    Dim sRegDate As String
    
    

On Error GoTo Err_p
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '������ ������� üũ
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    Set rs = New ADODB.Recordset
'    sQry = "SELECT Content FROM tb_config WHERE (NAME = 'WebDC' AND CONTENT = 'Y') "
'    rs.Open sQry, adoConn
'    If (rs.EOF) Then
'        Exit Sub
'        Set rs = Nothing
'    End If
'    Set rs = Nothing
    If (Glo_WebDC_YN = "N") Then
        Exit Sub
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    
    
    Call DataLogger("[Webdc Charge FreeAutoPoint]    " & "[������ �ڵ���������] �����췯 ����")
    sQry = "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('', 'HOST','�ڵ��������� �����췯 ����','SYSTEM'," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
    adoConn.Execute sQry
    
    
    sQry = "SELECT * FROM tb_id WHERE GUBUN != '�Ѱ�������' AND GUBUN != '������' AND GUBUN != '���' " '��� ��Ʈ��
    
    sRegDate = Format(Now, "yyyy-mm-dd hh:nn:ss")
    
    Set rs = New ADODB.Recordset
    rs.Open sQry, adoConn
    Do While Not (rs.EOF)
        
            sSEQ = "" & rs!SEQ
            sID = "" & rs!ID
            
            sQry = "SELECT * FROM tb_partner WHERE SEQ = '" & sSEQ & "'"
            Set rs2 = New ADODB.Recordset
            rs2.Open sQry, adoConn
            If Not (rs2.EOF) Then
            
                If (Left(rs2!FREE_AUTOPOINT_LASTDATE, 10) < Left(sRegDate, 10)) Then
                    sPName = "" & rs2!PNAME
                    nAutoFreePoint = rs2!FREE_AUTOPOINT
                    sLog = "[������ �ڵ���������]" & sSEQ & "." & sID & "(" & sPName & "):" & nAutoFreePoint & "(��)"
        
                    sQry = "UPDATE  tb_partner  SET  FREE_POINT = " & nAutoFreePoint & ", FREE_AUTOPOINT_LASTDATE = '" & sRegDate & "' WHERE SEQ = '" & sSEQ & "' "
                    adoConn.Execute sQry
                    
                    sQry = "INSERT INTO tb_partner_log (PCODE, FREE_POINT, INFO, CHARGE_ACCOUNT, REG_DATE) values ('" & sSEQ & "', " & nAutoFreePoint & ", '" & sLog & "', 'SYSTEM', '" & sRegDate & "' )"
                    adoConn.Execute sQry
                    
                    sQry = "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('" & sSEQ & "', 'HOST','" & sLog & "','SYSTEM'," & 0 & ",'" & sRegDate & "')"
                    adoConn.Execute sQry
                    
                    Call DataLogger("[WebDC Charge FreeAutoPoint]    " & sLog)
                End If
            Else
                Set rs2 = Nothing
            End If

        rs.MoveNext
    Loop

    Set rs2 = Nothing
    Set rs = Nothing
    
    Call DataLogger("[Webdc Charge FreeAutoPoint]    " & "[������ �ڵ���������] �����췯 ����")
    sQry = "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('', 'HOST','�ڵ��������� �����췯 ����','SYSTEM'," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
    adoConn.Execute sQry
    
    Exit Sub
    
Err_p:
    Set rs = Nothing
    Call DataLogger("[Webdc Charge FreeAutoPoint]    " & "�ڵ��������� �����ٷ� ����. �ٽ� �õ����ּ���(E00008)" & " " & Err.Description)
End Sub


'�ſ�1ȸ �湮Ƚ������ �������� ����Ƚ�� �ʱ�ȭ(�ſ� 1ȸ ����)
Public Sub Init_GuestInCarCount()
    
On Error GoTo Err_p
    
    If (Glo_GuestReg_YN = "N") Then
        Exit Sub
    End If
    
    Dim sQry As String
    Dim sLastUpdateDate As String
    Dim sNowDate As String
    Dim sLog As String
    
    sNowDate = Format(Now, "yyyy-mm-dd hh:nn:ss")
    
    sQry = "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('', 'HOST','[�����湮����] �ʱ�ȭ �����췯 ����(�湮Ƚ��)','SYSTEM'," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
    adoConn.Execute sQry
    
    Set rs = New ADODB.Recordset
    rs.Open "SELECT UPDATE_DATE FROM tb_guest_limit LIMIT 1", adoConn
    If Not (rs.EOF) Then
        sLastUpdateDate = Left(rs!UPDATE_DATE, 7)  'yyyy-mm
        
        If (sLastUpdateDate < Left(sNowDate, 7)) Then
            
            adoConn.Execute "UPDATE tb_guest_limit     SET NOWINPARK = 0"
            
            sLog = "[�����湮����]    " & " �湮���� ����Ƚ�� �ʱ�ȭ ���� "
            Call DataLogger(sLog)
            sQry = "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('', 'HOST','" & sLog & "','SYSTEM'," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
            adoConn.Execute sQry
        End If
    Else
            sLog = "[�����湮����]    " & " �湮���� ����Ƚ�� �ʱ�ȭ(������ ����) "
            Call DataLogger(sLog)
            sQry = "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('', 'HOST','" & sLog & "','SYSTEM'," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
            adoConn.Execute sQry
    End If
    
    Set rs = Nothing
    
    sQry = "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('', 'HOST','[�����湮����] �ʱ�ȭ �����췯 ����(�湮Ƚ��)','SYSTEM'," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
    adoConn.Execute sQry
    Exit Sub
    
Err_p:
    Set rs = Nothing
    Call DataLogger("[�����湮����]    " & "�ʱ�ȭ ����(E00008)" & " " & Err.Description)
    
End Sub

Private Sub Timer1_Timer()
    
    Dim qry As String
    Dim rs As ADODB.Recordset

On Error GoTo Err_p

    If (Format(Now, "NNSS") = "0001") Then
        '����Ʈ ī��Ʈ �ʱ�ȭ
    '    Qry = "show tables"
    '    Set rs = New ADODB.Recordset
    '    rs.Open Qry, adoConn
    '    Set rs = Nothing
        Call Time_Sync
        Call Inout_Reduce
        Call GuestLog_Backup
'        Call AppPW_Check '����Ͼ� ��й�ȣ �ʱ⺯�� ���� Ȯ��(�̺���� ����Ͼ� �̻����·� ����)
        Call Delete_GuestRegCar '�����湮���� ��������߿��� �Ⱓ�ʰ� ���� ����ó��
        Call Init_GuestRegCar '�ſ�1ȸ ��� ���� �����湮��û �Ǽ� �ʱ�ȭ
        Call Init_GuestInCarCount '�ſ�1ȸ �湮Ƚ������ �������� ����Ƚ�� �ʱ�ȭ
        Call Webdc_Charge_FreeAutoPoint '������, �ڵ�����(��������Ʈ)

    
'''    ElseIf (Format(Now, "NNSS") = "0030") Then
'''        '�Ϸ� �ѹ� �����湮���� ������ �����ð� ���(������, ��,ȣ����)
'''        If (Glo_GuestReg_YN = "Y") Then
'''            Call FrmGuestRegLog.GuestRegParkTime_Daily
'''        End If
    End If


    Call MainMessage(Not (DB_Connect_F), DB_Conn_Msg)

'''    Dim GCarno As String
'''    Dim serialCarNo As String * 4
'''    Dim i, no As Long
'''
'''    For i = 0 To 9999
'''        serialCarNo = Right("0000" & CStr(i), 4)
'''        GCarno = CStr(frontno) & "��" & serialCarNo
'''        adoConn.Execute "INSERT INTO tb_guest_log (GUEST_NO, CAR_NO, OBJECT, DONG, HO, HO2, NAME,TEL,ETC,ETC2,ETC3,DT_IN,IN_GATE,IN_IMAGE_PATH,DT_OUT,OUT_GATE,OUT_IMAGE_PATH,REG_DATE,DT_UPDATE,PARK_TIME ) VALUES ('', '" & GCarno & "','" & 0 & "','" & 0 & "','" & 0 & "', '" & 0 & "', '" & 0 & "','" & 10 & "','" & 0 & "','" & 0 & "','" & 0 & "','" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "','" & 0 & "','" & 0 & "','','','','" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "','', 0)"
'''
'''        If (i = 9999) Then
'''            frontno = frontno + 1
'''        End If
'''        If frontno = 100 Then
'''            Timer1.Enabled = False
'''        End If
'''    Next i

'''    For i = 0 To 9999
'''        serialCarNo = Right("0000" & CStr(i), 4)
'''        GCarno = CStr(frontno) & "��" & serialCarNo
'''        Call FormGuest1.Guest_Out_Auto_Proc(1, GCarno, Format(Now, "yyyy-mm-dd hh:nn:ss"), "", "�ⱸ")
'''
'''        If (i = 9999) Then
'''            frontno = frontno + 1
'''        End If
'''        If frontno = 100 Then
'''            Timer1.Enabled = False
'''        End If
'''    Next i

'''    Timer1.Enabled = False
'''    Dim bQryResult As Boolean
'''    Dim CarNo As String
'''    QRY = "SELECT car_no From tb_guest_log WHERE PARK_TIME = 0"
'''    Set rs = New ADODB.Recordset
'''    bQryResult = DataBaseQuery(rs, adoConn, QRY, False)
'''
'''    Do While Not (rs.EOF)
'''        CarNo = rs!CAR_NO
'''        Call FormGuest1.Guest_Out_Auto_Proc(1, CarNo, Format(Now, "yyyy-mm-dd hh:nn:ss"), "", "�ⱸ")
'''        rs.MoveNext
'''    Loop
'''
'''    Debug.Print "����ð�:" & Format(Now, "yyyy-mm-dd hh:nn:ss")
'''
'''    Set rs = Nothing
    
    
    

    
    
Exit Sub
Err_p:

End Sub

' 5�ʸ��� �ѹ��� üũ
Private Sub DBTimer_Timer()

    DBSock.Close
    DBSock.Protocol = sckTCPProtocol
    DBSock.Connect DB_Server_IP, DB_Server_Port
End Sub


Private Sub DBSock_DataArrival(ByVal bytesTotal As Long)

    Dim bRet As Boolean
    Dim sdata As String
    
    
On Error GoTo Err_p

    DBSock.GetData sdata, , bytesTotal
    'Debug.Print "Rcv Data:" & sdata & "(" & bytesTotal & " bytes)"
    
    
    If (sdata = "") Then
        Exit Sub
    End If
    
    'Debug.Print sdata
    ' DB�������� Connection �ź��� ���(err desc:Host is blocked because of many connection errors; Unblock with "mysqladmin flush-hosts")
    ' DB�������� flush hosts ������� Ŀ�ؼ� �ʱ�ȭ�ؾ� ��
    If (InStr(sdata, "block") > 0) Then
        Call DebugLogger("[DataBase] Connection failed:" & sdata)
        DB_Conn_Msg = "�����ͺ��̽� ������ ������ϴ�. Ȯ�ιٶ��ϴ�(DB_WARN-001)"
    
    ElseIf (InStr(sdata, "mysql") > 0) Then
    
        DB_Rcv_LastTime = Timer

        If (DB_Connect_F = False) Then
            Call DataBaseClose(adoConn)
            
            If (DataBaseOpen(adoConn) = True) Then
                DB_Connect_F = True
                'Debug.Print "DB ���� ���� ==> " & sdata
                Call DebugLogger("[DataBase] ReConnection successed")
                Exit Sub
            Else
                DB_Connect_F = False
                DB_Conn_Msg = "�����ͺ��̽� ������ ������ϴ�. Ȯ�ιٶ��ϴ�(DB_WARN-002)"
                'Call DataBaseClose(adoConn)
            End If
            
        Else
            DB_Conn_Msg = ""
        End If
    Else
        DB_Connect_F = False
        Call DebugLogger("[DataBase] Connection failed:" & sdata)
    End If
    
    
    
    
    If (Abs(Timer - DB_Rcv_LastTime) > 10) Then
        DB_Connect_F = False
    End If

    
    
'''    If (DB_Connect_F = False) Then
'''
'''        If (Glo_Screen_No = 6) Then
'''            FrmG6_23.LblDBInfo.Caption = "��Ʈ��ũ ������ ������ϴ�. Ȯ���� ��������ּ���."
'''            FrmG6_23.LblDBInfo.Visible = FrmG6_23.LblDBInfo.Visible Xor True
'''        ElseIf (Glo_Screen_No = 4) Then
'''            FrmG4Mini.LblDBInfo.Caption = "��Ʈ��ũ ������ ������ϴ�. Ȯ���� ��������ּ���."
'''            FrmG4Mini.LblDBInfo.Visible = FrmG4Mini.LblDBInfo.Visible Xor True
'''        ElseIf (Glo_Screen_No = 2) Then
'''            Jung.LblDBInfo.Caption = "��Ʈ��ũ ������ ������ϴ�. Ȯ���� ��������ּ���."
'''            Jung.LblDBInfo.Visible = Jung.LblDBInfo.Visible Xor True
'''        ElseIf (Glo_Screen_No = 1) Then
'''            FrmG1.LblDBInfo.Caption = "��Ʈ��ũ ������ ������ϴ�. Ȯ���� ��������ּ���."
'''            FrmG1.LblDBInfo.Visible = FrmG1.LblDBInfo.Visible Xor True
'''        End If
'''    Else
'''        If (Glo_Screen_No = 6) Then
'''            FrmG6_23.LblDBInfo.Visible = False
'''        ElseIf (Glo_Screen_No = 4) Then
'''            FrmG4Mini.LblDBInfo.Visible = False
'''        ElseIf (Glo_Screen_No = 2) Then
'''            Jung.LblDBInfo.Visible = False
'''        ElseIf (Glo_Screen_No = 1) Then
'''            FrmG1.LblDBInfo.Visible = False
'''        End If
'''    End If
    
    Exit Sub
Err_p:
    Call DebugLogger("[DataBase] Exception and resume : " & Err.Description)
    On Error Resume Next
    

End Sub
Private Sub DBSock_Connect()
'Debug.Print "DBSock_Connect"
End Sub
Private Sub DBSock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'Debug.Print "DBSock_Error"
DB_Connect_F = False
DB_Conn_Msg = "�����ͺ��̽� ������ ������ϴ�. Ȯ�ιٶ��ϴ�."
End Sub
Private Sub DBSock_Close()
'Debug.Print "DBSock_Close"
End Sub
Private Sub DBSock_ConnectionRequest(ByVal requestID As Long)
'Debug.Print "DBSock_ConnectionRequest"
End Sub

Public Sub MainMessage(bVisible As Boolean, msg As String)
    If (bVisible = True) Then
        If (Glo_Screen_No = 6) Then
            FrmG6_23.LblDBInfo.Caption = msg
            FrmG6_23.LblDBInfo.Visible = FrmG6_23.LblDBInfo.Visible Xor True
        ElseIf (Glo_Screen_No = 4) Then
            FrmG4Mini.LblDBInfo.Caption = msg
            FrmG4Mini.LblDBInfo.Visible = FrmG4Mini.LblDBInfo.Visible Xor True
        ElseIf (Glo_Screen_No = 2) Then
            Jung.LblDBInfo.Caption = msg
            Jung.LblDBInfo.Visible = Jung.LblDBInfo.Visible Xor True
        ElseIf (Glo_Screen_No = 1) Then
            FrmG1.LblDBInfo.Caption = msg
            FrmG1.LblDBInfo.Visible = FrmG1.LblDBInfo.Visible Xor True
        End If
    Else
        If (Glo_Screen_No = 6) Then
            FrmG6_23.LblDBInfo.Visible = False
        ElseIf (Glo_Screen_No = 4) Then
            FrmG4Mini.LblDBInfo.Visible = False
        ElseIf (Glo_Screen_No = 2) Then
            Jung.LblDBInfo.Visible = False
        ElseIf (Glo_Screen_No = 1) Then
            FrmG1.LblDBInfo.Visible = False
        End If
    
    End If
End Sub
Public Sub Inout_Reduce()

    Dim iDelDate As String

On Error GoTo Err_p
    
        If (Glo_INOUT_USING_DATE <> 99) Then
            
            iDelDate = DateAdd("m", Glo_INOUT_USING_DATE * (-1), Format(Now, "yyyy-mm-dd"))
            iDelDate = iDelDate + " 00:00:00 000"
            
            adoConn.Execute "INSERT INTO tb_inout_backup select * from tb_inout where PASS_DATE < '" & iDelDate & "' "
            adoConn.Execute "Delete from tb_inout WHERE PASS_DATE < '" & iDelDate & "'"
        
            Call DataLogger("������DB ��� : " & iDelDate & " ���� ��¥ ���")
    
        End If
Exit Sub

Err_p:
    Call DataLogger("Inout_Reduce Proc Error : " & Err.Description)
End Sub


Private Sub GuestLog_Backup()
    
    Dim iDelDate As String
    Dim sLog As String

On Error GoTo Err_p

    If (Glo_GuestLogBackup_YN = "Y") Then
        If (Glo_GuestLogBackup_Month > 0) Then

            ' �湮�� ������ �ֱ� Glo_GuestLogBackup_Month �ڷḸ ���ܵ�. �������� ���.
            iDelDate = DateAdd("m", Glo_GuestLogBackup_Month * (-1), Format(Now, "yyyy-mm-dd"))
            iDelDate = iDelDate + " 00:00:00"
            
            adoConn.Execute "INSERT INTO tb_guest_log_backup select * from tb_guest_log where REG_DATE < '" & iDelDate & "' "
            adoConn.Execute "Delete from tb_guest_log WHERE REG_DATE < '" & iDelDate & "'"
        
            sLog = "�湮���α׹��:" & iDelDate & " ���� ��¥ ���"
            adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('���', 'HOST','" & sLog & "',''," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
           'adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('', 'ȣ��Ʈ','ȣ��Ʈ ��Ʈ Open ����','" & PortNo & "'," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
            Call DataLogger("�湮��DB ��� : " & iDelDate & " ���� ��¥ ���")
        End If
    End If
Exit Sub
    
Err_p:
    Call DataLogger("Inout_Reduce Proc Error : " & Err.Description)
End Sub

Public Sub Disp1_sock_Connect(Index As Integer)
Select Case Index
    Case 0
        Disp1_sock(Index).SendData GloDisp_BData1
        Disp1_sock(Index).SendData GloDisp_BData1_Down
    Case 1
        Disp1_sock(Index).SendData GloDisp_BData2
        Disp1_sock(Index).SendData GloDisp_BData2_Down
    Case 2
        Disp1_sock(Index).SendData GloDisp_BData3
        Disp1_sock(Index).SendData GloDisp_BData3_Down
    Case 3
        Disp1_sock(Index).SendData GloDisp_BData4
        Disp1_sock(Index).SendData GloDisp_BData4_Down
    Case 4
        Disp1_sock(Index).SendData GloDisp_BData5
        Disp1_sock(Index).SendData GloDisp_BData5_Down

End Select
End Sub

Public Sub Disp1_sock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim strData As String
Dim bData() As Byte
Dim i As Integer
On Error Resume Next

Disp1_sock(Index).GetData bData, , bytesTotal - 1
'For i = 0 To (FormatNumber(bytesTotal, 0, , , vbTrue) - 1)
'    strData = strData & " " & Hex(bData(i))
'Next i
Call DataLogger("[Disp1 Rcv] " & FormatNumber(bytesTotal, 0, , , vbTrue) & " bytes recieved." & "    " & strData)
End Sub


Public Sub Gate1_sock_Connect(Index As Integer)
On Error GoTo Err_p
    Call DataLogger("[GATE TCP/IP ����]  �Ϸ� LANE" & Index + 1)
    Dim bData() As Byte
    ReDim bData(Len(GlO_TcpDataGate) - 1) As Byte
    bData = StrConv(GlO_TcpDataGate, vbFromUnicode)
    Gate1_sock(Index).SendData bData
    Exit Sub
Err_p:
    Call DataLogger("[GATE LANE" & Index & " Connect] " & " ���� : " & Err.Description)
    Call DebugLogger("[GATE LANE" & Index & " Connect] " & " ���� : " & Err.Description)
End Sub
Private Sub Gate1_sock_SendComplete(Index As Integer)
    Call DataLogger("[GATE TCP/IP ����]  �Ϸ� LANE" & Index + 1)
End Sub
Public Sub Gate1_sock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim strData As String
    Dim bData() As Byte
    Dim i As Integer

On Error GoTo Err_p

    'Gate1_sock(Index).GetData bData, , bytesTotal - 1
    
    Gate1_sock(Index).GetData strData, , bytesTotal
    
    
    If (Asc(strData) = 6) Then
        Gate_ACK(Index) = True
        Call DataLogger("[GATE LANE" & Index + 1 & " Rcv] " & "ACK")
    Else
        Call DataLogger("[GATE LANE" & Index + 1 & " Rcv] " & strData)
    End If
    
    
    'For i = 0 To (FormatNumber(bytesTotal, 0, , , vbTrue) - 1)
    '    strData = strData & " " & Hex(bData(i))
    'Next i
    
    
'''    Gate1_sock(Index).Close
    
    Exit Sub
    
Err_p:
    Call DataLogger("[LANE" & Index & " DataArrival] " & "���� : " & Err.Description)
    Call DebugLogger("[LANE" & Index & " DataArrival] " & "���� : " & Err.Description)
End Sub
Public Sub Gate1_sock_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    
On Error GoTo Err_p
    
    Call DataLogger("[GATE TCP/IP ����]  " & "LANE" & Index & " : " & Description)
    Call DebugLogger("[GATE TCP/IP ����]  " & "LANE" & Index & " : " & Description)
    
    Glo_Gate_ReconnCnt(Index) = Glo_Gate_ReconnCnt(Index) + 1
    If (Glo_Gate_ReconnCnt(Index) < 3) Then

        Gate1_sock(Index).Close
        Select Case Index
            Case 0
                Call DataLogger("[GATE TCP/IP ��õ�]  �õ� IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort)
                Call DebugLogger("[GATE TCP/IP ��õ�]  �õ� IP = " & LANE1_DeviceIP & "    PORT = " & LANE1_RelayPort)
                Gate1_sock(Index).Connect LANE1_DeviceIP, LANE1_RelayPort
            Case 1
                Call DataLogger("[GATE TCP/IP ��õ�]  �õ� IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort)
                Call DebugLogger("[GATE TCP/IP ��õ�]  �õ� IP = " & LANE2_DeviceIP & "    PORT = " & LANE2_RelayPort)
                Gate1_sock(Index).Connect LANE2_DeviceIP, LANE2_RelayPort
            Case 2
                Call DataLogger("[GATE TCP/IP ��õ�]  �õ� IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_RelayPort)
                Call DebugLogger("[GATE TCP/IP ��õ�]  �õ� IP = " & LANE3_DeviceIP & "    PORT = " & LANE3_RelayPort)
                Gate1_sock(Index).Connect LANE3_DeviceIP, LANE3_RelayPort
            Case 3
                Call DataLogger("[GATE TCP/IP ��õ�]  �õ� IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort)
                Call DebugLogger("[GATE TCP/IP ��õ�]  �õ� IP = " & LANE4_DeviceIP & "    PORT = " & LANE4_RelayPort)
                Gate1_sock(Index).Connect LANE4_DeviceIP, LANE4_RelayPort
            Case 4
                Call DataLogger("[GATE TCP/IP ��õ�]  �õ� IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_RelayPort)
                Call DebugLogger("[GATE TCP/IP ��õ�]  �õ� IP = " & LANE5_DeviceIP & "    PORT = " & LANE5_RelayPort)
                Gate1_sock(Index).Connect LANE5_DeviceIP, LANE5_RelayPort
            Case 5
                Call DataLogger("[GATE TCP/IP ��õ�]  �õ� IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_RelayPort)
                Call DebugLogger("[GATE TCP/IP ��õ�]  �õ� IP = " & LANE6_DeviceIP & "    PORT = " & LANE6_RelayPort)
                Gate1_sock(Index).Connect LANE6_DeviceIP, LANE6_RelayPort
        End Select
    End If
    
    Exit Sub
Err_p:
    Call DebugLogger("[GATE TCP/IP Err]  " & Err.Description)

End Sub



Public Sub LPR1_sock_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim sdata As String
    Dim Tmp_Path As String
    Dim i, gateNo As Integer
    Dim carnum As String
    Dim tmp_str As String
    
    
    
On Error GoTo Err_p
    
    
    If (bytesTotal > 500) Then
        'DebugLogger ("LPR ������ �ʰ�����(������) : " & bytesTotal)
        Exit Sub
    End If
    
    
    LPR1_sock(Index).GetData sdata, , bytesTotal
    Call DataLogger("-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------")
    '������ ��ȣȭ
    tmp_str = DecodeNDE01(sdata, "www.jawootek.com")
    'Call DataLogger("Lane" & Index + 1 & " UDP Port " & FormatNumber(bytesTotal, 0, , , vbTrue) & " bytes recieved." & "    " & tmp_str)
    Call DataLogger("Lane" & Left(tmp_str, 1) + 1 & " UDP Port " & FormatNumber(bytesTotal, 0, , , vbTrue) & " bytes recieved." & "    " & tmp_str)
    

    Select Case Mid(tmp_str, InStr(1, tmp_str, "Lane", 1), 5)
           Case "Lane1"
                LPR_Send_sock(0).RemoteHost = LPR1_sock(Index).RemoteHostIP
                LPR_Send_sock(0).RemotePort = 20101
                LPR_Send_sock(0).SendData "ACK"
           Case "Lane2"
                LPR_Send_sock(1).RemoteHost = LPR1_sock(Index).RemoteHostIP
                LPR_Send_sock(1).RemotePort = 20102
                LPR_Send_sock(1).SendData "ACK"
           Case "Lane3"
                LPR_Send_sock(2).RemoteHost = LPR1_sock(Index).RemoteHostIP
                LPR_Send_sock(2).RemotePort = 20103
                LPR_Send_sock(2).SendData "ACK"
           Case "Lane4"
                LPR_Send_sock(3).RemoteHost = LPR1_sock(Index).RemoteHostIP
                LPR_Send_sock(3).RemotePort = 20104
                LPR_Send_sock(3).SendData "ACK"
            Case "Lane5"
                LPR_Send_sock(4).RemoteHost = LPR1_sock(Index).RemoteHostIP
                LPR_Send_sock(4).RemotePort = 20105
                LPR_Send_sock(4).SendData "ACK"
            Case "Lane6"
                LPR_Send_sock(5).RemoteHost = LPR1_sock(Index).RemoteHostIP
                LPR_Send_sock(5).RemotePort = 20106
                LPR_Send_sock(5).SendData "ACK"
    End Select
    
    Call UDP_Proc(tmp_str)

Exit Sub

Err_p:
    Call DataLogger(" [Lane1 UDP DataArrival]  " & Err.Description)

End Sub

Public Sub LPR1_sock_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Call DataLogger(" [Lane" & Index + 1 & " UDP Error]  " & Description)
End Sub



'Reomte_UDP �ޱ�
Public Sub RemoteR_sock_DataArrival(ByVal bytesTotal As Long)
    Dim sdata As String
    Dim Tmp_Path As String
    Dim i, gateNo As Integer
    Dim carnum As String
    Dim sPassDate As String
    Dim sStrLine() As String
    
On Error GoTo Err_p

    If (bytesTotal > 500) Then
        'DebugLogger ("RemoteR ������ �ʰ�����(������) : " & bytesTotal)
        Exit Sub
    End If
    
    
    RemoteR_sock.GetData sdata, , bytesTotal
    sdata = "" & sdata
    
    If (sdata = "") Then
        Exit Sub
    End If
    
    sStrLine() = Split(sdata, "_")
    
'''    GateNo = Left(sdata, 1)
'''    i = Len(sdata)
'''    carnum = Mid(sdata, 3, i - 2)

    gateNo = sStrLine(0)
    carnum = sStrLine(1)
    
'    Debug.Print carnum
    
    Glo_Mon_LastInTime = Timer
    
    If (carnum = "LIVE") Then       ' ����͸� ����(ȣ��Ʈ���� ���� üũ)
        Glo_Mon_Lane(gateNo) = True
        Glo_MonStat_Lane(gateNo) = "LIVE"
        'Glo_Mon_LastInTime = Timer
    ElseIf (carnum = "DEAD") Then       ' ����͸� ����(ȣ��Ʈ���� ���� üũ)
        Glo_Mon_Lane(gateNo) = True
        Glo_MonStat_Lane(gateNo) = "DEAD"
        Call DataLogger("RemoteR Lane Stat : " & sdata)
        'Glo_Mon_LastInTime = Timer

     Else
         Call DataLogger("RemoteR_sock UDP Port " & FormatNumber(bytesTotal, 0, , , vbTrue) & " bytes recieved." & "    " & sdata)
         Glo_GateNo = gateNo
         sPassDate = sStrLine(2)
         
         
        Glo_Mon_Lane(gateNo) = True
        Glo_MonStat_Lane(gateNo) = "LIVE"
        
        
         '��ũ�� ���� ���� �б�
         If (Glo_Screen_No = 6) Then
             If (gateNo < Glo_Screen_No) Then
                 Call G6_23Show(carnum, gateNo, sPassDate)
             End If
         ElseIf (Glo_Screen_No = 4) Then
             If (gateNo < Glo_Screen_No) Then
                 Call G4Mini_4INShow(carnum, gateNo, sPassDate)
             End If
         ElseIf (Glo_Screen_No = 2) Then
             If (gateNo < Glo_Screen_No) Then
                 Call Jung_Show(carnum, gateNo, sPassDate)
             End If
         ElseIf (Glo_Screen_No = 1) Then
             If (gateNo < Glo_Screen_No) Then
                 Call G1_Show(carnum, gateNo, sPassDate)
             End If
         End If
     End If


    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '�����뺸
    If (HomeNet_YN = "Y") Then
        Dim rsR As ADODB.Recordset
        Dim qry As String
        
        Set rsR = New ADODB.Recordset
        qry = "SELECT * FROM tb_reg WHERE CAR_NO = '" & carnum & "'"
        'rs.Open Qry, adoConn
        If (DataBaseQuery(rsR, adoConn, qry, NWERR_GATE_STAY) = False) Then
            DataLogger ("[LPRIN_PROC]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�, ��������� �������_���ܱ� �ڵ� ����")
            Exit Sub
        End If

        If Not (rsR.EOF) Then
            If (IsNumeric(rsR!DRIVER_DEPT) = True) And (IsNumeric(rsR!DRIVER_CLASS) = True) And (rsR!DAY_ROTATION_YN = "����") Then
    
                HomeNet_Dong = rsR!DRIVER_DEPT
                HomeNet_Ho = rsR!DRIVER_CLASS
                HomeNet_CarNo = carnum
                
                HomeNet_Str = HomeNet_Dong & HomeNet_Ho & HomeNet_CarNo
                
                If (FrmTcpServer.HomeSock.State = sckClosed) Then
                    
                    FrmTcpServer.HomeSock.Protocol = sckUDPProtocol
                    FrmTcpServer.HomeSock.RemoteHost = HomeNet_IP
                    FrmTcpServer.HomeSock.RemotePort = HomeNet_Port
                    
                    FrmTcpServer.HomeSock.SendData (HomeNet_Str)
                    Call DataLogger("[HomeNet UDP ����]  DATA = " & HomeNet_Str)
                Else
                    FrmTcpServer.HomeSock.SendData (HomeNet_Str)
                    Call DataLogger("[HomeNet UDP ����]  DATA = " & HomeNet_Str)
                End If
            End If
        End If
        
        Set rsR = Nothing
    End If
    

Exit Sub

Err_p:
    Set rsR = Nothing
    Call DataLogger(" [RemoteR_sock UDP DataArrival]  " & Err.Description)

End Sub


Public Sub RemoteR_sock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Call DataLogger(" [RemoteR_sock UDP Error]  " & Description)
End Sub


Public Sub LPR_Alive_State_Send(ByVal iLaneNo As Integer, ByVal sState As String)
On Error GoTo Err_p

    If Glo_RemoteS_YN = "Y" Then
    
        Dim sSend_Str As String

        sSend_Str = CStr(iLaneNo) & "_" & sState & "_" & Format(Now, "yyyy-mm-dd hh:nn:ss") & Format(Timer * 1000 Mod 1000, " 000")
        'sSend_Str = CStr(iLaneNo + Glo_RemoteS_ScrPos) & "_" & sState & "_" & Format(Now, "yyyy-mm-dd hh:nn:ss") & Format(Timer * 1000 Mod 1000, " 000")

        RemoteS_sock.SendData (sSend_Str)
        'Call DataLogger("[LPR ���� ����]  DATA = " & sSend_Str)
    End If
    Exit Sub

Err_p:
    Call DataLogger(" [LPR Alive State]  " & Err.Description)
End Sub



Private Sub txt_CertifyKey_GotFocus()
    If (txt_CertifyKey.text = "����Ű �Է��ϼ���") Then
        txt_CertifyKey.text = ""
    End If
End Sub



Public Sub Print_Port_Init(Index As Integer, UseYN As String, Model As String, PortNo As String)

On Error Resume Next

    If MSComm(Index).PortOpen = True Then
        MSComm(Index).PortOpen = False
    End If
        
    If (UseYN = "N" Or Model = "NONE") Then
        Exit Sub
    End If
    
    
    'Model  : NONE, FILE, WRP-100P
    'PortNo : LPT1, LPT2, COM1~COM20
    If (PortNo = "LPT1") Then
    ElseIf (PortNo = "LPT2") Then
    ElseIf (InStr(1, PortNo, "COM") > 0) Then

        'If (Port_Init(MSComm(Index), UseYN, PortNo, 19200, 8, "n", 1) = True) Then
        If (Port_Init(MSComm(Index), UseYN, PortNo, 9600, 8, "n", 1) = True) Then
            Glo_Guest_Print_Open(Index) = "Y"
            Call DataLogger("LANE " & CStr(Index) & " ��Ʈ Open ���� : " & PortNo)
            'adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('', 'ȣ��Ʈ','ȣ��Ʈ ��Ʈ Open ����','" & PortNo & "'," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
            
        Else
            Glo_Guest_Print_Open(Index) = "N"
            Call DataLogger("LANE " & CStr(Index) & " ��Ʈ Open ���� : " & PortNo)
            'adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('', 'ȣ��Ʈ','������ ��Ʈ Open ����','" & PortNo & "'," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
        End If
    End If
    
    Glo_Guest_Print_Port(Index) = PortNo
        

    Exit Sub
    
Err_p:
    Glo_Guest_Print_Open(Index) = "N"
    Call DataLogger("ȣ��Ʈ ��Ʈ Open ���� : " & PortNo)
    'adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('', 'ȣ��Ʈ','������ ��Ʈ Open ����','" & PortNo & "'," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
End Sub



Public Function Port_Init(Port As MSComm, UseFlag As String, Optional PortNo As String, Optional Baud As Long, Optional DataBit As Integer, Optional ParityBit As String, Optional StopBit As Integer) As Boolean

    Dim port_Set As String
    Dim Ret As Integer

    Dim port_UseYN As String
    Dim port_Num As String
    Dim port_Parity As String
    Dim port_Baud As Long
    Dim port_Data As Integer
    Dim port_Stop As Integer

    
    port_UseYN = UseFlag
    port_Num = PortNo
    port_Baud = Baud
    port_Parity = ParityBit
    port_Data = DataBit
    port_Stop = StopBit

    If ((port_UseYN = "N") Or (Port.PortOpen = True)) Then
        Port_Init = True
        Exit Function
    End If
    
On Error GoTo Err_Proc

    Select Case port_Num
        
        Case "LPT1", "LPT2"
            Port_Init = True
            Exit Function
        Case Else
            Port.CommPort = CInt(Replace(PortNo, "COM", ""))

    End Select


    port_Set = port_Baud & "," & port_Parity & "," & port_Data & "," & port_Stop
    Port.Settings = port_Set
    
    Port.InputMode = comInputModeBinary
    Port.InputLen = 0
    Port.PortOpen = True
    
    Port_Init = True
    
Exit Function

Err_Proc:
    Port_Init = False
End Function



Private Sub txt_SiteName_Change()
    If (LenH(txt_SiteName) > 32) Then
        txt_SiteName = LeftH(txt_SiteName, 32)
    End If
End Sub
Private Sub txt_SiteName_LostFocus()
    txt_SiteName = lbl_SiteName
End Sub
Private Sub txt_SiteName_Validate(Cancel As Boolean)
    lbl_SiteName = txt_SiteName
End Sub

Private Sub txt_Vendor_Change()
    If (LenH(txt_Vendor) > 32) Then
        txt_Vendor = LeftH(txt_Vendor, 32)
    End If
End Sub
Private Sub txt_Vendor_LostFocus()
    txt_Vendor = lbl_Vendor
End Sub
Private Sub txt_Vendor_Validate(Cancel As Boolean)
    lbl_Vendor = txt_Vendor
End Sub





Private Sub Winsock_GateAgentR_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim sdata As String
    Dim sLog As String
    
On Error GoTo Err_p
    
    Winsock_GateAgentR(Index).GetData sdata, , bytesTotal

    sLog = "LANE" & Index + 1 & " : " & sdata
    adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('[GATE_AGENT]', 'HOST', '" & sLog & "',''," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
    Call DataLogger("[GATE_AGENT_R] " & sLog)
    Exit Sub
Err_p:
    Call DebugLogger("[Winsock_GateAgentR] DataArrival Error : " & Err.Description)
End Sub


Private Sub WinsockS_CertPC_Connect()
    On Error GoTo Err_p
    
    If (Len(WinsockS_CertPC) = 0) Then
        Exit Sub
    End If
    
    Dim bData() As Byte
    ReDim bData(Len(GlO_CertPC_TcpData) - 1) As Byte
    bData = StrConv(GlO_CertPC_TcpData, vbFromUnicode)
    WinsockS_CertPC.SendData bData
    
    'Call DataLogger("[��������]  �Ϸ�")
    Me.MousePointer = 0
    Exit Sub
Err_p:
    Me.MousePointer = 0
    Call DataLogger("[������] " & " �������� ����!! : " & Err.Description)
End Sub

Private Sub WinsockS_CertPC_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    Dim sStrLine() As String
    Dim iCount As Integer
    Dim iDataIndex As Integer
    Dim i As Integer
    Dim sLogStr As String
    Dim bSUCC As Boolean
    Dim sSiteCode As String
    Dim sSiteName As String
    Dim sVendorName As String

On Error GoTo Err_p

    WinsockS_CertPC.GetData strData, , bytesTotal
    

    strData = DecodeNDE01(strData, "www.jawootek.com")
    

    sStrLine() = Split(strData, "_")
    
    If (sStrLine(0) = "RES") Then
        If (sStrLine(1) = "VERIFY") Then
                If (sStrLine(2) = "SUCC") Then
                    sSiteCode = sStrLine(3)
                    sSiteName = sStrLine(4)
                    Call SaveClientKey(sSiteCode, sSiteName)
                    'Call DataLogger("[���üũ] ������ ȣ��ƮPC �Դϴ�")
                    
                ElseIf (sStrLine(2) = "NOREG") Then
                    Call DataLogger("[���üũ] �̵�� ȣ��ƮPC �Դϴ�. �ý��۰����ڿ��� �����ϼ���")
                    
                ElseIf (sStrLine(2) = "CERTWAIT") Then
                    Call DataLogger("[���üũ] ������� �����Դϴ�. �ý��۰����ڿ��� �����ϼ���")
                
                Else
                    Call DataLogger("[���üũ] �˼����� �����Դϴ�. �ý��۰����ڿ��� �����ϼ���")
                End If
                
        ElseIf (sStrLine(1) = "KEYREG") Then
            If (sStrLine(2) = "SUCC") Then
                'localhost DB �� �����ؾ� ��
                sSiteCode = sStrLine(3)
                sSiteName = sStrLine(4)
                'Public Function SaveClientKey(ByVal IP As String, ByVal Mac As String, ByVal Key As String, sSiteCode As String, sSiteName As String) As Boolean
                Call SaveClientKey(sSiteCode, sSiteName)
                Call DataLogger("[��ϼ���] ȣ��ƮPC ����߽��ϴ�")
                
            ElseIf (sStrLine(2) = "FAIL") Then
                
                If (sStrLine(3) = "DUP") Then
                    Call DataLogger("[��Ͻ���!!] �̹� ��ϵ� ȣ��ƮPC �Դϴ�")
                    sSiteCode = sStrLine(4)
                    sSiteName = sStrLine(5)
                    Call SaveClientKey(sSiteCode, sSiteName) '�������� ��ϵǾ� ������ Ŭ���̾�Ʈ���� �̵���� �� �����Ƿ� ���ó�� ��
                ElseIf (sStrLine(3) = "ERR") Then
                    Call DataLogger("[��Ͻ���!!] �� �� ���� ���� �Դϴ�. �ý��۰����ڿ��� �����ϼ���")
                    Call DeleteClientKey("FAIL")
                ElseIf (sStrLine(3) = "CERTWAIT") Then
                    Call DataLogger("[��Ͻ���!!] ������� �����Դϴ�. �ý��۰����ڿ��� �����ϼ���")
                    Call DeleteClientKey("WAIT")
                ElseIf (sStrLine(3) = "PARAM") Then
                    Call DataLogger("[��Ͻ���!!] �Ķ���� ���� �Դϴ�. �ý��۰����ڿ��� �����ϼ���")
                    Call DeleteClientKey("FAIL")
                End If
            End If
            
        ElseIf (sStrLine(1) = "SITEREG") Then
            
            If (sStrLine(2) = "SUCC") Then
                sVendorName = sStrLine(3)
                sSiteName = sStrLine(4)
                'Call SetSiteName(sSiteName)
                adoConn.Execute "UPDATE tb_certify   SET SITENAME = '" & sSiteName & "' "
                Call DataLogger("[��������] ����� �������߽��Դϴ�.")
                
            ElseIf (sStrLine(2) = "FAIL") Then
                Call DataLogger("[��������!!] ����� ��Ͻ����Դϴ�. �ý��۰����ڿ��� �����ϼ���")
                
            Else
                Call DataLogger("[��������!!] �� �� ���� ���� �Դϴ�. �ý��۰����ڿ��� �����ϼ���")
                
            End If
        End If
        
    End If
    
    
    'ListData.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & sLogStr, 0
    Exit Sub
    
Err_p:
    Call DebugLogger("[Socket DataArrival] " & "���� : " & Err.Description)
End Sub

Private Sub SetSiteNae(sSiteName As String)
    adoConn.Execute "UPDATE tb_certify SET SITENAME='" & sSiteName & "' "
End Sub

Private Sub WinsockS_CertPC_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    On Error GoTo Err_p

    'Call DataLogger("[�������� ����]  " & Description)
    'Call DataLogger("[�ý��� �����ڿ��� ���ǹٶ�]")
        
    WinsockS_CertPC.Close

    Me.MousePointer = 0
    Exit Sub
Err_p:
    Call DataLogger("[������] " & " ��� ����(���Ͽ���) : " & Err.Description)
End Sub



Private Sub Certify_PC()
    Dim sQry As String
    Dim bQryResult As Boolean
    Dim rs As Recordset
    
    Dim sVendorName As String
    Dim sSiteName As String
    
    
On Error GoTo Err_p

    Dim rSiteCode As String
    Dim rSiteName As String
    
    Call GetClientIP(Glo_IPAddr)
    Call GetClientMac(Glo_MacAddr)
    Call GetClienKey(Glo_PhyHDDKey)
    
    'DB���� PC���� ���ٸ� ����
    'PC������ �������� ���� Ȯ����
    Call SaveCertPC(Glo_IPAddr, Glo_MacAddr, Glo_PhyHDDKey)
    
    
    
    '���� ���ؼ� �����޾ƾ� ��
    'Call DebugLogger("[CERTIFY SERVER CONNECTING...]")
    Glo_Certify_PC = False
    If (GetCertServerInfo(Glo_CertServerIP, Glo_CertServerPORT) = False) Then '����IP, PORT ȹ��
        Exit Sub
    End If
    
'    Call Certify_PC_Process '���μ������� ����(���� ������. �����ʿ�)
    

    Exit Sub
Err_p:

End Sub

Private Sub Certify_PC_Process()

On Error GoTo Err_p
    
    Call SendCertPacket("REQ_KEYREG_" & Glo_IPAddr & "_" & Glo_MacAddr & "_" & Glo_PhyHDDKey & "_" & "ȣ��Ʈ")
    
    
'    Dim rs As Recordset
'    Dim Qry As String
'
'    Qry = "SELECT * From tb_certify WHERE HASHCODE = '" & UniqKey & "' "
'    Set rs = New ADODB.Recordset
'    rs.Open Qry, adoConn
'
'    Do While Not (rs.EOF)
'        If (Len(Trim(rs!ip)) > 0 And Len(Trim(rs!mac)) > 0 And Len(Trim(rs!HASHCODE)) > 0) Then
'            Call SendCertPacket("REQ_KEYREG_" & Trim(rs!ip) & "_" & Trim(rs!mac) & "_" & Trim(rs!HASHCODE) & "_" & "ȣ��Ʈ")
'        Else
'
'        End If
'    End If
    Exit Sub
Err_p:
    Call DebugLogger("[Certify_PC_Process] " & Err.Description)
    
End Sub

Private Sub SendCertPacket(ByVal sdata As String)
    Dim ECHO As ICMP_ECHO_REPLY
    Dim RemoteIP As String
    Call Ping(Glo_CertServerIP, ECHO)
    If Left$(ECHO.Data, 1) <> Chr$(0) Then
        
        GlO_CertPC_TcpData = EncodeNDE01(sdata, "www.jawootek.com")
        FrmTcpServer.WinsockS_CertPC.Close
        FrmTcpServer.WinsockS_CertPC.Connect Glo_CertServerIP, Glo_CertServerPORT
        
        'Call DataLogger("[������]  ���� ������...")
    Else
        'Call DataLogger("[������]  ���� ���� ����!!")
    End If
End Sub


'ȣ��ƮPC�� ����DB�� ��ϾȵǾ� �ִٸ� �����
Private Sub SaveCertPC(ip As String, mac As String, UniqKey As String)
    Dim rs As Recordset
    Dim qry As String
    
    On Error GoTo Err_p
    
    qry = "SELECT * From tb_certify WHERE HASHCODE = '" & UniqKey & "' "
    Set rs = New ADODB.Recordset
    rs.Open qry, adoConn
    
    If (rs.EOF) Then
        adoConn.Execute "INSERT INTO tb_certify (IP, MAC, HASHCODE, SITECODE, SITENAME, C2DATE) VALUE ('" & ip & "', '" & mac & "', '" & UniqKey & "', '000000', '������', '" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "') "
    End If
    Set rs = Nothing
    
    Exit Sub
Err_p:
    Set rs = Nothing
End Sub


'Ŭ���̾�Ʈ ���� �ܺ�IP
'����:Microsoft WinHTTP Services, version 5.1
Public Sub GetClientIP(sClientIP As String)
    Dim WinHttp As Object
    Set WinHttp = CreateObject("Winhttp.WinHttpRequest.5.1")

    WinHttp.Open "GET", "http://map.naver.com"
    WinHttp.Send
    WinHttp.WaitForResponse: DoEvents
    If (InStr(WinHttp.ResponseText, "userIP")) Then
        sClientIP = Split(Split(WinHttp.ResponseText, "userIP")(1), """")(2)
    Else
        'sClientIP = "ERROR:���ͳ� ������ �������� �ʽ��ϴ�. �ٽ� �õ����ּ���"
        sClientIP = ""
    End If
End Sub
'Ŭ���̾�Ʈ �ƾ�巹��
Public Sub GetClientMac(sClientMac As String)
    Dim ls_ConnectIP As String
    Dim ls_MacAddress As String
    Dim ls_PcName As String
    
On Error GoTo Err_p
    sClientMac = ""
    
    ls_ConnectIP = Space(255)
    ls_PcName = Space(255)

    GetIPAddress ls_ConnectIP, 128
    ls_ConnectIP = Left(ls_ConnectIP, InStr(ls_ConnectIP, Chr(0)) - 1)
    GetComputerName ls_PcName, 128
    ls_PcName = Left(ls_PcName, InStr(ls_PcName, Chr(0)) - 1)
    ls_MacAddress = Gf_MACAddress
    
    If (Len(ls_MacAddress) > 0) Then
        sClientMac = ls_MacAddress
    Else
        sClientMac = "ERROR:�ƾ�巹�� ���� �Դϴ�."
    End If

    Exit Sub
    
Err_p:
    Call DataLogger("Get MacAddress Err: " & Err.Description)
End Sub
'Ű����
Public Sub GetClienKey(sKey As String)
    Dim msg As String

    On Error Resume Next
    
    msg = GetHDDID
    msg = EncodeNDE01(msg, "www.jawootek.com")
    'msg = GetCPUID
    'Call GetClientMac(msg)
    
    
    If (Len(msg) > 0) Then
        sKey = msg
    Else
        sKey = "Ű�� ȹ�����!!"
    End If
End Sub
Public Function GetHDDID() As String
    Dim WMI As Variant
    Dim PhysicalMedia As Variant
    Dim Media As Variant
    Dim MediaSerial As String
    
    Set WMI = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & "." & "\root\cimv2")
    Set PhysicalMedia = WMI.ExecQuery("Select * from " & "Win32_PhysicalMedia")

    For Each Media In PhysicalMedia
        MediaSerial = Media.SerialNumber
        Exit For
    Next Media
    
    GetHDDID = Trim(MediaSerial)
End Function




'ȣ��ƮPC ������ ��������
Private Function GetCertServerInfo(rSvrIP As String, rSvrPort As Long) As Boolean
    
    On Error GoTo ErrorHandler

    Dim pHandle     As Long
    Dim lRet        As Long
    Dim strShellCommand As String
    Dim lngTask As Long
    
    Dim LF          As Long
    Dim strLine     As String
    Dim sStrData()  As String

    GetCertServerInfo = False
    rSvrIP = ""
    rSvrPort = 0
    
    strShellCommand = "cmd.exe /c ping -n 1 jawootek.iptime.org > log.txt "
    lngTask = Shell(strShellCommand, vbHide)
    If lngTask <> 0 Then
        pHandle = OpenProcess(SYNCHRONIZE, 0, lngTask)
        Do
            lRet = WaitForSingleObject(pHandle, INFINITE)
            DoEvents
        Loop While lRet <> 0
        
        
        LF = FreeFile()
        Open "log.txt" For Input As LF
        
        Do While Not EOF(LF)
            Line Input #LF, strLine
            If (InStr(strLine, "jawootek.iptime.org") > 0) Then
                Exit Do
            End If
            
        Loop
        Close #LF
    Else
        Call DataLogger("Command ���� �Դϴ�. �ý��� �����ڿ��� �����ϼ���(20011)")
    End If
    
    
    If (Len(strLine) > 0) Then
        sStrData() = Split(strLine, "[")
        sStrData() = Split(sStrData(1), "]")
        rSvrIP = sStrData(0)
        rSvrPort = 35000
        
        GetCertServerInfo = True
    Else
        rSvrIP = ""
        rSvrPort = 0
        Call DataLogger("�������� ȹ�� ���� �Դϴ�. �ý��� �����ڿ��� �����ϼ���(20012)")

    End If
    
    strShellCommand = "cmd.exe /c del log.txt "
    lngTask = Shell(strShellCommand, vbHide)
    Exit Function

ErrorHandler:
    Call DataLogger("�ý��� ���� �Դϴ�. �ý��� �����ڿ��� �����ϼ���(20013)")
    Close #LF
    
End Function



