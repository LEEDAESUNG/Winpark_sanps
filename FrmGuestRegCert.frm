VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMctl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmGuestRegCert 
   BackColor       =   &H00404040&
   BorderStyle     =   1  '���� ����
   Caption         =   "ParkingManager��"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15840
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   15840
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   " �˻� / ���� "
      BeginProperty Font 
         Name            =   "�������"
         Size            =   11.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   3510
      Left            =   120
      TabIndex        =   3
      Top             =   4800
      Width           =   15585
      Begin VB.TextBox txt_NowParkCount 
         Height          =   315
         Left            =   13755
         TabIndex        =   38
         Text            =   "txt_NowParkCount"
         Top             =   3075
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.TextBox txt_MaxGuestVisitCount_Default 
         BackColor       =   &H80000002&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6135
         TabIndex        =   36
         ToolTipText     =   "0 �Է½� �̻��"
         Top             =   2895
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.TextBox txt_MaxGuestVisitCount 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6135
         TabIndex        =   34
         ToolTipText     =   "0 �Է½� �̻��"
         Top             =   1740
         Width           =   1755
      End
      Begin VB.TextBox txt_MaxParkDay 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   31
         ToolTipText     =   "0 �Է½� �̻��"
         Top             =   1740
         Width           =   1755
      End
      Begin VB.TextBox txt_MaxParkDay_Default 
         BackColor       =   &H80000002&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   28
         ToolTipText     =   "0 �Է½� �̻��"
         Top             =   2895
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.ComboBox cmb_Dong 
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
         Left            =   375
         TabIndex        =   27
         Text            =   "cmb_Dong"
         Top             =   825
         Width           =   1155
      End
      Begin VB.TextBox txt_MaxParkTime_Default 
         BackColor       =   &H80000002&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4215
         TabIndex        =   23
         ToolTipText     =   "0 �Է½� �̻��"
         Top             =   2895
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.TextBox txt_MaxParkCount_Default 
         BackColor       =   &H80000002&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2295
         TabIndex        =   22
         ToolTipText     =   "0 �Է½� �̻��"
         Top             =   2895
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.TextBox txt_MaxParkCount 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2295
         TabIndex        =   16
         ToolTipText     =   "0 �Է½� �̻��"
         Top             =   1740
         Width           =   1755
      End
      Begin VB.TextBox txt_MaxParkTime 
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4215
         TabIndex        =   12
         ToolTipText     =   "0 �Է½� �̻��"
         Top             =   1740
         Width           =   1755
      End
      Begin VB.ComboBox cmb_Cert 
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
         Left            =   3000
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   9
         Top             =   825
         Width           =   1530
      End
      Begin VB.ComboBox cmb_Ho 
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
         Left            =   1695
         TabIndex        =   1
         Text            =   "cmb_Ho"
         Top             =   825
         Width           =   1155
      End
      Begin Threed.SSCommand cmd_Search 
         Height          =   690
         Left            =   10410
         TabIndex        =   6
         Top             =   285
         Width           =   1665
         _Version        =   65536
         _ExtentX        =   2937
         _ExtentY        =   1217
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
         Picture         =   "FrmGuestRegCert.frx":0000
      End
      Begin Threed.SSCommand SSCommand1 
         Height          =   690
         Left            =   12090
         TabIndex        =   7
         Top             =   285
         Width           =   1665
         _Version        =   65536
         _ExtentX        =   2937
         _ExtentY        =   1217
         _StockProps     =   78
         Caption         =   "����"
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
         Picture         =   "FrmGuestRegCert.frx":0351
      End
      Begin Threed.SSCommand SSCommand3 
         Height          =   690
         Left            =   13770
         TabIndex        =   11
         Top             =   285
         Width           =   1665
         _Version        =   65536
         _ExtentX        =   2937
         _ExtentY        =   1217
         _StockProps     =   78
         Caption         =   "����"
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
         Picture         =   "FrmGuestRegCert.frx":06A2
      End
      Begin Threed.SSCommand SSCommand4 
         Height          =   690
         Left            =   13770
         TabIndex        =   13
         ToolTipText     =   "�� ����� �湮�� �������� �����ð��հ踦 ������ ������ �����մϴ�"
         Top             =   1815
         Width           =   1665
         _Version        =   65536
         _ExtentX        =   2937
         _ExtentY        =   1217
         _StockProps     =   78
         Caption         =   "��ü�ð�����"
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
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "FrmGuestRegCert.frx":09F3
      End
      Begin Threed.SSCommand SSCommand5 
         Height          =   690
         Left            =   13770
         TabIndex        =   15
         ToolTipText     =   "��й�ȣ�� ""0000""���� �ʱ�ȭ�մϴ�."
         Top             =   1050
         Width           =   1665
         _Version        =   65536
         _ExtentX        =   2937
         _ExtentY        =   1217
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
         Picture         =   "FrmGuestRegCert.frx":0D44
      End
      Begin Threed.SSCommand SSCommand6 
         Height          =   690
         Left            =   12090
         TabIndex        =   18
         ToolTipText     =   "�����湮��û�� �ִ� ��ϰǼ�(��)�� ������ ������ ��� ���뿡 �����մϴ�"
         Top             =   1815
         Width           =   1665
         _Version        =   65536
         _ExtentX        =   2937
         _ExtentY        =   1217
         _StockProps     =   78
         Caption         =   "��ü�Ǽ�����"
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
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "FrmGuestRegCert.frx":1095
      End
      Begin Threed.SSCommand SSCommand8 
         Height          =   690
         Left            =   10410
         TabIndex        =   19
         ToolTipText     =   "�̽��� ������ ��� ���̵� ���Ͽ� ������ ó���մϴ�."
         Top             =   1050
         Width           =   1665
         _Version        =   65536
         _ExtentX        =   2937
         _ExtentY        =   1217
         _StockProps     =   78
         Caption         =   "��ü���Խ���"
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
         Picture         =   "FrmGuestRegCert.frx":13E6
      End
      Begin Threed.SSCommand SSCommand9 
         Height          =   690
         Left            =   12090
         TabIndex        =   21
         ToolTipText     =   "�����湮��û �α��� ID �ڵ������մϴ�. ����� ��/ȣ�� �������� �ű� �����Դϴ�."
         Top             =   1050
         Width           =   1665
         _Version        =   65536
         _ExtentX        =   2937
         _ExtentY        =   1217
         _StockProps     =   78
         Caption         =   "ID �ڵ�����"
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
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "FrmGuestRegCert.frx":1737
      End
      Begin Threed.SSCommand SSCommand10 
         Height          =   690
         Left            =   8055
         TabIndex        =   26
         ToolTipText     =   "�����ϼ�, ��ϰǼ�, �����ð� �⺻���� �����մϴ�. ���� ID�ڵ����� �Ǵ� ���� ���Խ�û ���뿡 ����˴ϴ�."
         Top             =   2580
         Visible         =   0   'False
         Width           =   1665
         _Version        =   65536
         _ExtentX        =   2937
         _ExtentY        =   1217
         _StockProps     =   78
         Caption         =   "�⺻������"
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
         Enabled         =   0   'False
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "FrmGuestRegCert.frx":1A88
      End
      Begin Threed.SSCommand SSCommand11 
         Height          =   690
         Left            =   10410
         TabIndex        =   30
         ToolTipText     =   "�����湮��û�� �ִ� �����ϼ�(������)�� ������ ������ ��� ���뿡 �����մϴ�"
         Top             =   1815
         Width           =   1665
         _Version        =   65536
         _ExtentX        =   2937
         _ExtentY        =   1217
         _StockProps     =   78
         Caption         =   "��ü�ϼ�����"
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
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "FrmGuestRegCert.frx":1DD9
      End
      Begin Threed.SSCommand SSCommand12 
         Height          =   690
         Left            =   10410
         TabIndex        =   33
         ToolTipText     =   "�����湮��û�� ��� ������ �����ð��հ�(��)�� ������ ������ ��� ���뿡 �����մϴ�"
         Top             =   2580
         Visible         =   0   'False
         Width           =   1665
         _Version        =   65536
         _ExtentX        =   2937
         _ExtentY        =   1217
         _StockProps     =   78
         Caption         =   "��ü�湮Ƚ��"
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
         Enabled         =   0   'False
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "FrmGuestRegCert.frx":212A
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "�湮Ƚ��(��)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   6135
         TabIndex        =   37
         Top             =   2550
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "�湮Ƚ��(��)"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   6135
         TabIndex        =   35
         Top             =   1395
         Width           =   1305
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "�����ϼ�(������)"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   360
         TabIndex        =   32
         Top             =   1395
         Width           =   1755
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "�����ϼ�(������)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   360
         TabIndex        =   29
         Top             =   2550
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000000&
         X1              =   360
         X2              =   10200
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "�����ð�(��)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   4215
         TabIndex        =   25
         Top             =   2550
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "��ϰǼ�(��)"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   2295
         TabIndex        =   24
         Top             =   2550
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Label lbl_parktcount 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "��ϰǼ�(��)"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   2295
         TabIndex        =   17
         Top             =   1395
         Width           =   1305
      End
      Begin VB.Label lbl_parktime 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "�����ð�(��)"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   4215
         TabIndex        =   14
         Top             =   1395
         Width           =   1305
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "���Խ��ο���"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   3000
         TabIndex        =   10
         Top             =   465
         Width           =   1350
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   375
         TabIndex        =   5
         Top             =   450
         Width           =   270
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "ȣ��"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   1695
         TabIndex        =   4
         Top             =   450
         Width           =   540
      End
   End
   Begin ComctlLib.ListView ListView_GuestRegCar 
      Height          =   3675
      Left            =   120
      TabIndex        =   0
      Top             =   930
      Width           =   15585
      _ExtentX        =   27490
      _ExtentY        =   6482
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9150
      Top             =   150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Threed.SSCommand SSCommand2 
      Cancel          =   -1  'True
      Height          =   570
      Left            =   14430
      TabIndex        =   8
      Top             =   105
      Width           =   1260
      _Version        =   65536
      _ExtentX        =   2222
      _ExtentY        =   1005
      _StockProps     =   78
      Caption         =   "�ݱ�"
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
      Picture         =   "FrmGuestRegCert.frx":247B
   End
   Begin Threed.SSCommand SSCommand7 
      Height          =   570
      Left            =   13020
      TabIndex        =   20
      Top             =   105
      Width           =   1260
      _Version        =   65536
      _ExtentX        =   2222
      _ExtentY        =   1005
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
      Picture         =   "FrmGuestRegCert.frx":27CC
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   0
      X1              =   135
      X2              =   15675
      Y1              =   795
      Y2              =   795
   End
   Begin VB.Label lbl_APS 
      BackStyle       =   0  '����
      Caption         =   "�����湮��û ���Խ��� / ID�ڵ�����"
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
      Height          =   300
      Index           =   0
      Left            =   165
      TabIndex        =   2
      Top             =   300
      Width           =   4470
   End
End
Attribute VB_Name = "FrmGuestRegCert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sLv_SelectID As String
Dim sLv_SelectDong As String
Dim sLv_SelectHo As String
Dim sLV_SelectCert As String
Dim sLV_SelectParkDay As String
Dim sLV_SelectParkTime As String
Dim sLV_SelectParkCount As String
Dim sLV_SelectNowParkCount As String
Dim sLV_SelectVisitCount As String

Const DEF_CERTIFY_USE As String = "����"
Const DEF_CERTIFY_NOTUSE As String = "�̽���"
Const DEF_INIT_PASSWORD As String = "0000" '��й�ȣ �ʱⰪ

Private Sub Form_Load()

    Left = (Screen.width - width) / 2   ' ���� ���η� �߾ӿ� �����ϴ�.
    Top = (Screen.height - height) / 2   ' ���� ���η� �߾ӿ� �����ϴ�.
    
    Call Clear_Field
    Call ListView_GuestRegCar_Draw
    Call ListView_GuestRegCar_SQL("SELECT * From tb_guestReg_admin ")
    
    Call Getconfig
    
End Sub

Private Sub Getconfig()

    Dim rs As Recordset
    Dim sQry As String
    Dim nMaxCount As Integer
    Dim nMaxTime As Integer
    Dim nMaxDay As Integer
    Dim nMaxVisitCount As Integer
    
    nMaxCount = 0
    nMaxTime = 0
    nMaxDay = 0
    
    sQry = "SELECT * FROM tb_config WHERE (NAME = 'GuestCarReg_MaxParkCount' OR NAME = 'GuestCarReg_MaxParkTime' OR NAME = 'GuestCarReg_MaxParkDay' OR NAME = 'GuestCarReg_MaxGuestVisitCount') "
    Set rs = New ADODB.Recordset
    rs.Open sQry, adoConn
    Do While Not (rs.EOF)
        If (rs!name = "GuestCarReg_MaxParkCount") Then
            nMaxCount = rs!Content
        End If
        If (rs!name = "GuestCarReg_MaxParkTime") Then
            nMaxTime = rs!Content
        End If
        If (rs!name = "GuestCarReg_MaxParkDay") Then
            nMaxDay = rs!Content
        End If
        If (rs!name = "GuestCarReg_MaxGuestVisitCount") Then
            nMaxVisitCount = rs!Content
        End If
        
        rs.MoveNext
    Loop
    
    txt_MaxParkCount_Default = nMaxCount
    txt_MaxParkTime_Default = nMaxTime
    txt_MaxParkDay_Default = nMaxDay
    txt_MaxGuestVisitCount_Default = nMaxVisitCount
    
End Sub
'����
Private Sub SSCommand1_Click()
    Dim sLog As String
    Dim sUse As String

On Error Resume Next
    
    If (sLv_SelectID = "") Then
        Exit Sub
    End If
    
    If (cmb_Cert.text = DEF_CERTIFY_USE) Then
        sUse = "Y"
    Else
        sUse = "N"
    End If

    sLog = "�湮���� ��������(��:" & sLv_SelectDong & "->" & cmb_Dong.text & ", ȣ:" & sLv_SelectHo & "->" & cmb_Ho.text & ", ������:" & sLV_SelectCert & "->" & cmb_Cert.text & ", �����ð�(��):" & sLV_SelectParkTime & "->" & txt_MaxParkTime & ", �����ִ�Ƚ��(��):" & sLV_SelectParkCount & "->" & txt_MaxParkCount & ", �����Ǽ�(��):" & sLV_SelectNowParkCount & "->" & txt_NowParkCount & ")"
    'adoConn.Execute "UPDATE tb_guestReg_admin SET DRIVER_DEPT = '" & cmb_Dong.text & "', DRIVER_CLASS = '" & cmb_Ho.text & "', USE_YN = '" & sUse & "', MAXPARKTIME = '" & txt_MaxParkTime & "', MAXPARKCOUNT = '" & txt_MaxParkCount & "'  WHERE ID = '" & sLv_SelectID & "' "
    adoConn.Execute "UPDATE tb_guestReg_admin SET DRIVER_DEPT = '" & cmb_Dong.text & "', DRIVER_CLASS = '" & cmb_Ho.text & "', USE_YN = '" & sUse & "', MAXPARKDAY = '" & txt_MaxParkDay & "', MAXPARKTIME = '" & txt_MaxParkTime & "', MAXPARKCOUNT = '" & txt_MaxParkCount & "', NOWPARKCOUNT = '" & txt_NowParkCount & "' WHERE DRIVER_DEPT = '" & sLv_SelectDong & "' AND DRIVER_CLASS = '" & sLv_SelectHo & "' "
    adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('�湮����', 'HOST', '" & sLog & "', '" & Glo_Login_ID & "', " & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
    Call DataLogger(sLog)

    Call cmd_Search_Click
End Sub


'�⺻�� ����(���� ID�ڵ����� �Ǵ� �������� ȸ������ �����)
' key :GuestCarReg_MaxParkCount / GuestCarReg_MaxParkTim / GuestCarReg_MaxParkDay / GuestCarReg_MaxGuestVisitCount
Private Sub SaveDefault(key As String, value As String)
    On Error Resume Next
    adoConn.Execute "UPDATE tb_config SET Content = '" & value & "' WHERE Name = '" & key & "' "
End Sub

Private Sub SSCommand10_Click()
'    Dim sLog As String
'    Dim sUse As String
'    Dim bCheck As Boolean
'
'On Error Resume Next
'
'
'    bCheck = True
'    If (IsNumeric(txt_MaxParkCount_Default) = False Or txt_MaxParkCount_Default < 0) Then
'        bCheck = False
'        txt_MaxParkCount_Default = "0"
'    End If
'    If (IsNumeric(txt_MaxParkTime_Default) = False Or txt_MaxParkTime_Default < 0) Then
'        bCheck = False
'        txt_MaxParkTime_Default = "0"
'    End If
'    If (IsNumeric(txt_MaxParkDay_Default) = False Or txt_MaxParkDay_Default < 0) Then
'        bCheck = False
'        txt_MaxParkDay_Default = "0"
'    End If
'    If (IsNumeric(txt_MaxGuestVisitCount_Default) = False Or txt_MaxGuestVisitCount_Default < 0) Then
'        bCheck = False
'        txt_MaxGuestVisitCount_Default = "0"
'    End If
'
'    If (bCheck = False) Then
'        Msg_Box.Label2.Caption = "�����湮���� - �⺻�� ����"
'        Msg_Box.Label1.Caption = "�⺻�� ���� �����Դϴ�." & vbCrLf & vbCrLf & "�缳�� �� �����ϼ���."
'        Msg_Box.Show 1
'        Exit Sub
'    End If
'
'    sLog = "�湮���� �⺻�� ����(�ִ������Ǽ�(��):" & txt_MaxParkCount_Default & ", �ִ������ð�(��):" & txt_MaxParkTime_Default & ", �ִ������Ⱓ(��):" & txt_MaxParkDay_Default
'    sLog = sLog
'    adoConn.Execute "UPDATE tb_config SET Content = '" & txt_MaxParkCount_Default & "' WHERE Name = 'GuestCarReg_MaxParkCount' "
'    adoConn.Execute "UPDATE tb_config SET Content = '" & txt_MaxParkTime_Default & "' WHERE Name = 'GuestCarReg_MaxParkTim' "
'    adoConn.Execute "UPDATE tb_config SET Content = '" & txt_MaxParkDay_Default & "' WHERE Name = 'GuestCarReg_MaxParkDay' "
'    adoConn.Execute "UPDATE tb_config SET Content = '" & txt_MaxGuestVisitCount_Default & "' WHERE Name = 'GuestCarReg_MaxGuestVisitCount' "
'
'    adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('�湮����', 'HOST', '" & sLog & "', '" & Glo_Login_ID & "', " & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
'    Call DataLogger(sLog & ", �ִ�湮Ƚ��(��):" & txt_MaxGuestVisitCount_Default)
'
'    Msg_Box.Label2.Caption = "�����湮���� - �⺻�� ����"
'    Msg_Box.Label1.Caption = "�⺻�� ���� �Ϸ��߽��ϴ�."
'    Msg_Box.Show 1

End Sub

'��� ���� ����
'�����湮���� ������ȣ ��Ͻ� �ִ� �����ϼ� ����
Private Sub SSCommand11_Click()
    txt_MaxParkDay = Trim(txt_MaxParkDay)
    
    If (sLv_SelectID = "") Then
        Msg_Box.Label2.Caption = "�����Ⱓ �ϰ�����"
        Msg_Box.Label1.Caption = "�׸��� �����ϼ���."
        Msg_Box.Show 1
        Exit Sub
    End If
    
    If (Val(txt_MaxParkDay) <= 0) Then
        Msg_Box.Label2.Caption = "�����Ⱓ �ϰ�����"
        Msg_Box.Label1.Caption = "�����Ⱓ�� �ùٸ��� �Է¹ٶ��ϴ�."
        Msg_Box.Show 1
        Exit Sub
    End If
    
    Dim sLog As String
    
    sLog = "[�����湮����]�ϰ����� ��ü�����ϼ�(������):" & txt_MaxParkDay & "(��)"
    adoConn.Execute "UPDATE tb_guestReg_admin SET MAXPARKDAY = '" & txt_MaxParkDay & "' "
    adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('�湮����', 'HOST', '" & sLog & "', 'Glo_Login_ID', " & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
    
    Call SaveDefault("GuestCarReg_MaxParkDay", txt_MaxParkDay) 'tb_config ����(�ű�ID ���� �Ǵ� ����ȸ�� ���Խ� �ʿ���)
    
    Call DataLogger(sLog)
    
    Call Clear_Field
    Call cmd_Search_Click
End Sub

''�湮Ƚ�� �ϰ�����
Private Sub SSCommand12_Click()
'    txt_MaxGuestVisitCount = Trim(txt_MaxGuestVisitCount)
'
'    If (sLv_SelectID = "") Then
'        Msg_Box.Label2.Caption = "�湮Ƚ�� �ϰ�����"
'        Msg_Box.Label1.Caption = "�׸��� �����ϼ���."
'        Msg_Box.Show 1
'        Exit Sub
'    End If
'
'    If (Val(txt_MaxGuestVisitCount) <= 0) Then
'        Msg_Box.Label2.Caption = "�湮Ƚ�� �ϰ�����"
'        Msg_Box.Label1.Caption = "�湮Ƚ�� �ùٸ��� �Է¹ٶ��ϴ�."
'        Msg_Box.Show 1
'        Exit Sub
'    End If
'
'    Dim sLog As String
'
'    sLog = "�ִ�湮Ƚ��(��) �ϰ�����:" & txt_MaxGuestVisitCount & "(ȸ)"
'    adoConn.Execute "UPDATE tb_guestReg_admin SET MAXGUESTVISITCOUNT = '" & txt_MaxGuestVisitCount & "' "
'    adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('�湮����', 'HOST', '" & sLog & "', 'Glo_Login_ID', " & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
'
'    Call SaveDefault("GuestCarReg_MaxGuestVisitCount", txt_MaxGuestVisitCount) 'tb_config ����(�ű�ID ���� �Ǵ� ����ȸ�� ���Խ� �ʿ���)
'
'    Call DataLogger(sLog)
'
'    Call Clear_Field
'    Call cmd_Search_Click
End Sub

'����
Private Sub SSCommand3_Click()
    If (sLv_SelectID = "") Then
        Msg_Box.Label2.Caption = "��������"
        Msg_Box.Label1.Caption = "������ �׸��� �����ϼ���."
        Msg_Box.Show 1
        Exit Sub
    End If
    
    
    Dim sID As String
    Dim sLog As String
    Dim i As Long
    Dim iListCount As Long
    
    iListCount = 0
    For i = 1 To ListView_GuestRegCar.ListItems.Count
        If ListView_GuestRegCar.ListItems(i).Selected = True Then
            iListCount = iListCount + 1
        End If
    Next i
    If (iListCount = 1) Then
        MBox.Label3.Caption = sLv_SelectID
    ElseIf (iListCount >= 2) Then
        MBox.Label3.Caption = sLv_SelectID & " �� " & iListCount - 1 & "��"
    End If
    MBox.Label1.Caption = "�ش� �׸��� �����մϴ�." & vbCrLf & vbCrLf & " �����Ͻðڽ��ϱ�?"
    MBox.Label2.Caption = "���̵����"
    MBox.Show 1
    If (Glo_MsgRet = True) Then
        For i = 1 To ListView_GuestRegCar.ListItems.Count
            If ListView_GuestRegCar.ListItems(i).Selected = True Then

                sID = ListView_GuestRegCar.ListItems(i).SubItems(3) '���̵�

                sLog = "�湮���� ���Խ�û ���̵����:" & sID
                adoConn.Execute "DELETE FROM tb_guestReg_admin WHERE ID = '" & sID & "'"
                adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('�湮����', 'HOST', '" & sLog & "', '" & Glo_Login_ID & "', " & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
                Call DataLogger(sLog)
            
            End If
        Next i
    End If
    
    Call Clear_Field
    Call cmd_Search_Click
End Sub



Private Sub SSCommand2_Click()
    Unload Me
    'Me.Hide
End Sub

Private Sub Clear_Field()
 
    SSCommand1.Enabled = False  '����
    SSCommand3.Enabled = False  '����
    'SSCommand4.Enabled = False  '�ð��ϰ�����
    SSCommand5.Enabled = False  '��й�ȣ �ʱ�ȭ
    'SSCommand6.Enabled = False  'Ƚ���ϰ�����
    sLv_SelectID = ""
    sLv_SelectDong = ""
    sLv_SelectHo = ""

    Call Set_cmbDong
    Call Set_cmbHo
    Call Set_cmbCert
End Sub

Private Sub ListView_GuestRegCar_Draw()
    Dim Column_to_size As Integer
    
    With Me
        Call ListViewExtended(.ListView_GuestRegCar)
        ListView_GuestRegCar.MultiSelect = True
        
        .ListView_GuestRegCar.View = lvwReport
        .ListView_GuestRegCar.ListItems.Clear
        .ListView_GuestRegCar.ColumnHeaders.Clear
        .ListView_GuestRegCar.ColumnHeaders.Add , , " No   "
        .ListView_GuestRegCar.ColumnHeaders.Add , , " ��            "
        .ListView_GuestRegCar.ColumnHeaders.Add , , " ȣ            "
        .ListView_GuestRegCar.ColumnHeaders.Add , , " ���̵�               "
        .ListView_GuestRegCar.ColumnHeaders.Add , , " ������ȣ             "
        .ListView_GuestRegCar.ColumnHeaders.Add , , " �̸�          "
        .ListView_GuestRegCar.ColumnHeaders.Add , , " ����ó               "
        .ListView_GuestRegCar.ColumnHeaders.Add , , " ���Խ��ο��� "
        .ListView_GuestRegCar.ColumnHeaders.Add , , " �ִ������ϼ�(������) "
        
        .ListView_GuestRegCar.ColumnHeaders.Add , , " �ִ������Ǽ�(��) "
        .ListView_GuestRegCar.ColumnHeaders.Add , , " ���������Ǽ� "
        
        .ListView_GuestRegCar.ColumnHeaders.Add , , " �ִ������ð�(��) "
        .ListView_GuestRegCar.ColumnHeaders.Add , , " ���������ð� "
        
        .ListView_GuestRegCar.ColumnHeaders.Add , , " �ִ�湮Ƚ��(��) "
        .ListView_GuestRegCar.ColumnHeaders.Add , , " ����湮Ƚ�� "
        
        .ListView_GuestRegCar.ColumnHeaders.Add , , " ���Խ�û�Ͻ�                 "
        .ListView_GuestRegCar.ColumnHeaders.Add , , ""
        
        For Column_to_size = 0 To .ListView_GuestRegCar.ColumnHeaders.Count - 2
             SendMessage .ListView_GuestRegCar.hwnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
        Next
    End With

End Sub


Private Sub ListView_GuestRegCar_SQL(qry As String)
    Dim bQryResult As Boolean
    Dim rs As Recordset
'    Dim qry As String
    Dim itmX As ListItem
    Dim INDEX_NO As Long
    Dim i As Integer
    
    
    Set rs = New ADODB.Recordset
    bQryResult = DataBaseQuery(rs, adoConn, qry, False)
    If (bQryResult = False) Then
        Exit Sub
    End If

    INDEX_NO = 1
    Do While Not (rs.EOF)
        Set itmX = ListView_GuestRegCar.ListItems.Add(, , "" & INDEX_NO)
        
        i = 1
        itmX.SubItems(i) = "" & rs!DRIVER_DEPT: i = i + 1   '��
        itmX.SubItems(i) = "" & rs!DRIVER_CLASS: i = i + 1  'ȣ��
        itmX.SubItems(i) = "" & rs!ID: i = i + 1            '���̵�
        itmX.SubItems(i) = "" & rs!carno: i = i + 1         '������ȣ
        itmX.SubItems(i) = "" & rs!name: i = i + 1          '�̸�
        itmX.SubItems(i) = "" & rs!Tel: i = i + 1           '����ó
        itmX.SubItems(i) = "" & rs!USE_YN: i = i + 1        '���Խ��ο���
        itmX.SubItems(i) = "" & rs!maxparkday: i = i + 1    '�����ϼ�(������)
        itmX.SubItems(i) = "" & rs!maxparkcount: i = i + 1  '�ִ������Ǽ�(��)
        itmX.SubItems(i) = "" & rs!nowparkcount: i = i + 1  '���������Ǽ�
        itmX.SubItems(i) = "" & rs!maxparktime: i = i + 1   '�ִ������ð�(��)
        itmX.SubItems(i) = "" & rs!nowparktime: i = i + 1   '�ִ������ð�(��)
        itmX.SubItems(i) = "" & rs!MAXGUESTVISITCOUNT: i = i + 1   '�ش缼����� �ִ�湮ȸ��(��)
        itmX.SubItems(i) = "" & rs!NOWGUESTVISITCOUNT: i = i + 1   '�ش缼����� ����湮ȸ��(ȸ)
        
        itmX.SubItems(i) = "" & rs!REG_DATE: i = i + 1      '���Խ�û�Ͻ�

        rs.MoveNext
        INDEX_NO = INDEX_NO + 1
    Loop
    Set rs = Nothing
End Sub


Private Sub ListView_GuestRegCar_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
    Dim i As Integer
    With ListView_GuestRegCar
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

Private Sub ListView_GuestRegCar_ItemClick(ByVal Item As ComctlLib.ListItem)
    On Error Resume Next
    
    SSCommand1.Enabled = True   '����
    SSCommand3.Enabled = True   '����
    SSCommand4.Enabled = True  '�ð��ϰ�����
    SSCommand5.Enabled = True   '��й�ȣ �ʱ�ȭ
    SSCommand6.Enabled = True  'Ƚ���ϰ�����
    
    ListView_GuestRegCar.SetFocus

    sLv_SelectID = ListView_GuestRegCar.SelectedItem.SubItems(3) '���̵�
    sLv_SelectDong = ListView_GuestRegCar.SelectedItem.SubItems(1) '��
    sLv_SelectHo = ListView_GuestRegCar.SelectedItem.SubItems(2) 'ȣ��
    sLV_SelectCert = ListView_GuestRegCar.SelectedItem.SubItems(7) '���Խ��� ����
    sLV_SelectParkDay = ListView_GuestRegCar.SelectedItem.SubItems(8) '�ִ������ϼ�(������)
    sLV_SelectParkCount = ListView_GuestRegCar.SelectedItem.SubItems(9) '�� �ִ������Ǽ�
    sLV_SelectNowParkCount = ListView_GuestRegCar.SelectedItem.SubItems(10) '������� �����Ǽ�(��)
    sLV_SelectParkTime = ListView_GuestRegCar.SelectedItem.SubItems(11) '�� �ִ������ð�
    'sLV_SelectParkTime = ListView_GuestRegCar.SelectedItem.SubItems(12) '������� ���������ð�
    sLV_SelectVisitCount = ListView_GuestRegCar.SelectedItem.SubItems(13) '�� �ִ�湮Ƚ��
    'sLV_SelectVisitCount = ListView_GuestRegCar.SelectedItem.SubItems(14) '��������湮Ƚ��
    
    cmb_Dong.text = ListView_GuestRegCar.SelectedItem.SubItems(1) '��
    cmb_Ho.text = ListView_GuestRegCar.SelectedItem.SubItems(2) 'ȣ��cmb_Cert
    If (ListView_GuestRegCar.SelectedItem.SubItems(7) = "Y") Then '���Խ��� ����
        cmb_Cert.text = DEF_CERTIFY_USE
    Else
        cmb_Cert.text = DEF_CERTIFY_NOTUSE
    End If
    txt_MaxParkDay.text = sLV_SelectParkDay
    txt_MaxParkTime.text = sLV_SelectParkTime
    txt_MaxParkCount.text = sLV_SelectParkCount
    txt_NowParkCount.text = sLV_SelectNowParkCount
    txt_MaxGuestVisitCount = sLV_SelectVisitCount
    
End Sub

Private Sub cmd_Search_Click()
    Dim sDong, sHo As String
    Dim sQry As String
    
    sDong = Trim(cmb_Dong.text)
    sHo = Trim(cmb_Ho.text)
    sQry = "SELECT * From tb_guestReg_admin "
    If (cmb_Dong.text = "��ü") Then
        If (cmb_Ho.text = "��ü") Then
        Else
            sQry = sQry & " WHERE DRIVER_CLASS = '" & cmb_Ho.text & "' "
        End If
    Else
        If (cmb_Ho.text = "��ü") Then
            sQry = sQry & " WHERE DRIVER_DEPT = '" & cmb_Dong.text & "' "
        Else
            sQry = sQry & " WHERE DRIVER_DEPT = '" & cmb_Dong.text & "' AND DRIVER_CLASS = '" & cmb_Ho.text & "' "
        End If
    End If
    
    If (cmb_Cert = "��ü") Then
    Else
        If (cmb_Cert = DEF_CERTIFY_USE) Then
            If (InStr(1, UCase(sQry), "WHERE") > 0) Then
                sQry = sQry & " AND USE_YN = 'Y' "
            Else
                sQry = sQry & " WHERE USE_YN = 'Y' "
            End If
        ElseIf (cmb_Cert = DEF_CERTIFY_NOTUSE) Then
            If (InStr(1, UCase(sQry), "WHERE") > 0) Then
                sQry = sQry & " AND USE_YN <> 'Y' "
            Else
                sQry = sQry & " WHERE USE_YN <> 'Y' "
            End If
        End If
    End If
    
    sQry = sQry & " ORDER BY DRIVER_DEPT, DRIVER_CLASS "
    
    Call Clear_Field
    Call ListView_GuestRegCar_Draw
    Call ListView_GuestRegCar_SQL(sQry)
End Sub


Private Sub Set_cmbDong()
    Dim bQryResult As Boolean
    Dim rs As Recordset
    Dim qry As String
    Dim nCount As Integer
On Error GoTo Err_p

    qry = "SELECT DRIVER_DEPT From tb_guestReg_admin Group By DRIVER_DEPT ORDER BY DRIVER_DEPT"

    Set rs = New ADODB.Recordset
     bQryResult = DataBaseQuery(rs, adoConn, qry, False)
     If (bQryResult = False) Then
        Call DataLogger("[FrmReg]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
        Exit Sub
    End If
    
    cmb_Dong.Clear
    cmb_Dong.AddItem "��ü"
    nCount = rs.RecordCount
    Do While Not (rs.EOF)
        cmb_Dong.AddItem "" & rs!DRIVER_DEPT
        rs.MoveNext
    Loop
    Set rs = Nothing
    
    If (nCount > 0) Then
        cmb_Dong.ListIndex = 0
    End If

Exit Sub
Err_p:
    Call DataLogger("[FrmGuestRegCert Set_cmbDong]    " & Err.Description)
End Sub

Private Sub Set_cmbHo()
    Dim bQryResult As Boolean
    Dim rs As Recordset
    Dim qry As String
    Dim nCount As Integer
On Error GoTo Err_p
    
    qry = "SELECT DRIVER_CLASS From tb_guestReg_admin Group By DRIVER_CLASS ORDER BY DRIVER_CLASS"
    
    Set rs = New ADODB.Recordset
     bQryResult = DataBaseQuery(rs, adoConn, qry, False)
     If (bQryResult = False) Then
        Call DataLogger("[FrmReg]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
        Exit Sub
    End If
    
    cmb_Ho.Clear
    cmb_Ho.AddItem "��ü"
    nCount = rs.RecordCount
    Do While Not (rs.EOF)
        cmb_Ho.AddItem "" & rs!DRIVER_CLASS
        rs.MoveNext
    Loop
    Set rs = Nothing
    
    If (nCount > 0) Then
        cmb_Ho.ListIndex = 0
    End If
Exit Sub

Err_p:
    Call DataLogger("[FrmGuestRegCert Set_cmbHo]    " & Err.Description)
End Sub

Private Sub Set_cmbCert()

On Error GoTo Err_p
    
    cmb_Cert.Clear
    cmb_Cert.AddItem "��ü"
    cmb_Cert.AddItem DEF_CERTIFY_USE
    cmb_Cert.AddItem DEF_CERTIFY_NOTUSE
    
    cmb_Cert.ListIndex = 0
Exit Sub

Err_p:
    Call DataLogger("[FrmGuestRegCert Set_cmbCert]    " & Err.Description)
End Sub

Private Sub SSCommand7_Click()
'    Dim tmpFileName As String
'    tmpFileName = Format(Now, "YYYYMMDD_HHMMSS")
'    tmpFileName = App.Path & "\Excel\" & tmpFileName & "_�湮���� ������� ����" & ".xls"
'    Call MakeCSV(ListView_GuestRegCar, tmpFileName)
    
    
    Dim tmpFileName As String
On Error GoTo Err_p
    tmpFileName = Format(Now, "YYYYMMDD_HHMMSS")
    tmpFileName = App.Path & "\Excel\" & tmpFileName & "_�湮���� �����ð�����"
        
        
    CommonDialog1.CancelError = True
    CommonDialog1.InitDir = App.Path
    CommonDialog1.Filter = "��������(*.csv)|*.csv"
    CommonDialog1.fileName = tmpFileName
    CommonDialog1.ShowSave
    tmpFileName = CommonDialog1.fileName
    tmpFileName = Mid(tmpFileName, 1, Len(tmpFileName) - 4)

    Call MakeCSV(ListView_GuestRegCar, tmpFileName)
    Exit Sub
Err_p:
     Select Case Err
    Case 32755 '  Dialog Cancelled
        'MsgBox "you cancelled the dialog box"
    Case Else
        'MsgBox "Unexpected error. Err " & Err & " : " & Error
    End Select
End Sub

'�����ð��ϰ�����
Private Sub SSCommand4_Click()

    txt_MaxParkTime = Trim(txt_MaxParkTime)
    
    If (sLv_SelectID = "") Then
        Msg_Box.Label2.Caption = "�����ð� �ϰ�����"
        Msg_Box.Label1.Caption = "�׸��� �����ϼ���."
        Msg_Box.Show 1
        Exit Sub
    End If
    
    If (Val(txt_MaxParkTime) <= 0) Then
        Msg_Box.Label2.Caption = "�����ð� �ϰ�����"
        Msg_Box.Label1.Caption = "�����ð��� �ùٸ��� �Է¹ٶ��ϴ�."
        Msg_Box.Show 1
        Exit Sub
    End If
    
    Dim sLog As String
    
    sLog = "�ִ������ð�(��) �ϰ�����:" & txt_MaxParkTime & "(��)"
    adoConn.Execute "UPDATE tb_guestReg_admin SET MAXPARKTIME = '" & txt_MaxParkTime & "' "
    adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('�湮����', 'HOST', '" & sLog & "', 'Glo_Login_ID', " & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
    
    Call SaveDefault("GuestCarReg_MaxParkTim", txt_MaxParkTime) 'tb_config ����(�ű�ID ���� �Ǵ� ����ȸ�� ���Խ� �ʿ���)
    
    Call DataLogger(sLog)
    
    Call Clear_Field
    Call cmd_Search_Click
End Sub


'�����׸� �ʱ�ȭ
'1.�����׸� ��/ȣ�������� ����ǿ��� �˻��Ͽ� �̸�,��ȭ��ȣ,������ȣ �����ͼ� ����
'2.�����׸� ��й�ȣ �ʱ�ȭ(0000)
'3.�����׸� �����湮��û �ִ�Ǽ� �ʱ�ȭ, tb_config:GuestCarReg_MaxParkCount(�����湮��û �ִ�Ǽ�)
Private Sub SSCommand5_Click()
    
    Dim rs1 As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim sQry As String
    Dim sRegCarno As String
    Dim sRegName As String
    Dim sRegTel As String
    Dim nParkCount As Integer
    Dim nParkTime As Integer
    Dim nParkDay As String
    Dim sTmpHo As String
    
    
    If (sLv_SelectID = "") Then
        Msg_Box.Label2.Caption = "[�����湮��û]"
        'Msg_Box.Label1.Caption = "�ʱ�ȭ�� �׸��� �����ϼ���."
        Msg_Box.Label1.Caption = "�׸��� �����ϼ���."
        Msg_Box.Show 1
        Exit Sub
    End If
    
    MBox.Label2.Caption = "[�����湮��û]"
    MBox.Label3.Caption = "ID:" & sLv_SelectID
    'MBox.Label1.Caption = "ȸ�������� �ʱ�ȭ �Ͻðڽ��ϱ�?"
    MBox.Label1.Caption = "��й�ȣ�� �ʱ�ȭ �Ͻðڽ��ϱ�?"
    MBox.Show 1
    
    If (Glo_MsgRet = True) Then
        Dim sLog As String
        
        'sLog = "ȸ������ �ʱ�ȭ ����:" & sLv_SelectID
        'adoConn.Execute "UPDATE tb_guestReg_admin SET PASSWORD = '" & DEF_INIT_PASSWORD & "' WHERE ID = '" & sLv_SelectID & "' "
        'adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('�湮����', 'HOST', '" & sLog & "', 'Glo_Login_ID', " & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
        'Call DataLogger(sLog)
        
        
        
'        '�⺻ ȯ�漳���� ��������
'        nParkCount = 0
'        nParkTime = 0
'        nParkDay = 0
'        sQry = "SELECT * FROM tb_config WHERE Name = 'GuestCarReg_MaxParkCount' OR Name = 'GuestCarReg_MaxParkTime' OR Name = 'GuestCarReg_MaxParkDay' " '��/ȣ�� ���� ��������
'        Set rs1 = New ADODB.Recordset
'        rs1.Open sQry, adoConn
'        Do While Not (rs1.EOF)
'
'            If (rs1!name = "GuestCarReg_MaxParkCount") Then     '�����湮��û �ִ��ûȽ��(��, �ſ�)
'                nParkCount = rs1!Content
'            ElseIf (rs1!name = "GuestCarReg_MaxParkTime") Then  '�����湮��û �ִ������ð�(��, �ſ�)
'                nParkTime = rs1!Content
'            ElseIf (rs1!name = "GuestCarReg_MaxParkDay") Then  '�����湮��û �ִ������Ⱓ(��)
'                nParkDay = rs1!Content
'            End If
'
'            rs1.MoveNext
'        Loop
'        Set rs1 = Nothing
        
        
        
        
'        sQry = "SELECT * FROM tb_reg WHERE (DRIVER_DEPT = '" & sLv_SelectDong & "' AND DRIVER_CLASS = '" & sLv_SelectHo & "') " '��/ȣ�� ���� ��������
'        Set rs2 = New ADODB.Recordset
'        rs2.Open sQry, adoConn
'        If Not (rs2.EOF) Then
'            sRegCarno = rs2!CAR_NO
'            sRegName = rs2!DRIVER_NAME
'            sRegTel = rs2!DRIVER_PHONE
'
'            adoConn.Execute "UPDATE tb_guestReg_admin SET PASSWORD = '" & DEF_INIT_PASSWORD & "', CARNO = '" & sRegCarno & "', NAME = '" & sRegName & "', TEL = '" & sRegTel & "', MAXPARKTIME = '" & nParkTime & "', MAXPARKCOUNT = " & nParkCount & ", MAXPARKDAY = '" & nParkDay & "', NOWPARKTIME = 0, NOWPARKCOUNT = 0, USE_YN = 'Y' WHERE ID = '" & sLv_SelectID & "' "
'            adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('�湮����', 'HOST', '" & sLog & "', '" & Glo_Login_ID & "', " & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
'
'            Msg_Box.Label2.Caption = "�ʱ�ȭ"
'            Msg_Box.Label1.Caption = "��й�ȣ '0000'�� �ʱ�ȭ�߽��ϴ�."
'            Msg_Box.Show 1
'
'            sLog = "��й�ȣ/ȸ������ �ʱ�ȭ ����:" & sLv_SelectID
'            adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('�湮����', 'HOST', '" & sLog & "', 'Glo_Login_ID', " & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
'            Call DataLogger(sLog)
'
'        Else
'
'            sLog = "��й�ȣ/ȸ������ �ʱ�ȭ ����(����� ������ ����):" & sLv_SelectID
'            adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('�湮����', 'HOST', '" & sLog & "', 'Glo_Login_ID', " & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
'            Call DataLogger(sLog)
'
''            Msg_Box.Label2.Caption = "�ʱ�ȭ ����"
''            Msg_Box.Label1.Caption = "����ǿ� ��/ȣ���� �����ϴ�"
''            Msg_Box.Show 1
'
'        End If
'        Set rs2 = Nothing

        
        sLog = "��й�ȣ �ʱ�ȭ:" & sLv_SelectID
        adoConn.Execute "UPDATE tb_guestReg_admin SET PASSWORD = '" & DEF_INIT_PASSWORD & "' "
        adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('�湮����', 'HOST', '" & sLog & "', '" & Glo_Login_ID & "', " & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
        Call DataLogger(sLog)
        Msg_Box.Label2.Caption = "��й�ȣ �ʱ�ȭ"
        Msg_Box.Label1.Caption = "��й�ȣ '0000'���� �ʱ�ȭ�߽��ϴ�."
        Msg_Box.Show 1

        
        Call Clear_Field
        Call cmd_Search_Click
    End If
    
End Sub

'�����Ǽ��ϰ�����
Private Sub SSCommand6_Click()
    txt_MaxParkCount = Trim(txt_MaxParkCount)
    
    If (sLv_SelectID = "") Then
        Msg_Box.Label2.Caption = "�����Ǽ� �ϰ�����"
        Msg_Box.Label1.Caption = "�׸��� �����ϼ���."
        Msg_Box.Show 1
        Exit Sub
    End If
    
    If (Val(txt_MaxParkCount) <= 0) Then
        Msg_Box.Label2.Caption = "�����Ǽ� �ϰ�����"
        Msg_Box.Label1.Caption = "�����Ǽ��� �ùٸ��� �Է¹ٶ��ϴ�."
        Msg_Box.Show 1
        Exit Sub
    End If
    
    Dim sLog As String
    
    sLog = "[�����湮����]�ϰ����� �ִ������Ǽ�(��):" & txt_MaxParkCount & "(��)" & ", �����Ǽ�(��):" & txt_NowParkCount & "(��)"
    adoConn.Execute "UPDATE tb_guestReg_admin SET MAXPARKCount = '" & txt_MaxParkCount & "', NOWPARKCOUNT = '" & txt_NowParkCount & "' "
    adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('�湮����', 'HOST', '" & sLog & "', 'Glo_Login_ID', " & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
    
    Call SaveDefault("GuestCarReg_MaxParkCount", txt_MaxParkCount) 'tb_config ����(�ű�ID ���� �Ǵ� ����ȸ�� ���Խ� �ʿ���)
    
    Call DataLogger(sLog)
    
    Call Clear_Field
    Call cmd_Search_Click
End Sub

'��ü���Խ���
Private Sub SSCommand8_Click()
    
    MBox.Label3.Caption = "��ü���Խ���"
    MBox.Label1.Caption = "��� ���Խ�û ����ó���մϴ�." & vbCrLf & " �����Ͻðڽ��ϱ�?"
    MBox.Label2.Caption = "�湮����"
    MBox.Show 1
    If (Glo_MsgRet = True) Then
    
        Dim sLog As String
        
        sLog = "��ü���Խ��� �ϰ�ó��"
        adoConn.Execute "UPDATE tb_guestReg_admin SET USE_YN = 'Y' "
        adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('�湮����', 'HOST', '" & sLog & "', 'Glo_Login_ID', " & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
        Call DataLogger(sLog)
        
        Call Clear_Field
        Call cmd_Search_Click
    End If
End Sub


' ����� ��/ȣ�� ��� ���̵� �ڵ� ����
' ��(3�ڸ�), ȣ��(4�ڸ�) �� ������ ID�� �ű� �����Ѵ�.
' ���� ID�� ���� ��� pass
Private Sub SSCommand9_Click()
    Dim rs1 As Recordset
    Dim rs2 As Recordset
    Dim sQry As String
    Dim sDong As String
    Dim sHo As String
    Dim sID As String
    Dim nCount As Integer
    Dim nParkCount As Integer
    Dim nParkTime As Integer
    Dim nParkDay As Integer
    Dim nVisitCount As Integer
    
    
On Error GoTo Err_p
    

    Call DataLogger("[�����湮��û] ID �ڵ� ���� ����")
    
    nParkCount = 0
    nParkTime = 0
    sQry = "SELECT * FROM tb_config WHERE Name = 'GuestCarReg_MaxParkCount' or Name = 'GuestCarReg_MaxParkTime' or Name = 'GuestCarReg_MaxParkDay'" '��/ȣ�� ���� ��������
    Set rs1 = New ADODB.Recordset
    rs1.Open sQry, adoConn
    
    Do While Not (rs1.EOF)
        
        If (rs1!name = "GuestCarReg_MaxParkCount") Then     '�����湮��û �ִ��ûȽ��(��, �ſ�)
            nParkCount = rs1!Content
        ElseIf (rs1!name = "GuestCarReg_MaxParkTime") Then  '�����湮��û �ִ������ð�(��, �ſ�)
            nParkTime = rs1!Content
        ElseIf (rs1!name = "GuestCarReg_MaxParkDay") Then   '�����湮��û �ִ������Ⱓ(��)
            nParkDay = rs1!Content
        ElseIf (rs1!name = "GuestCarReg_MaxParkDay") Then   '����湮 �ִ�湮Ƚ��(ȸ)
            nVisitCount = rs1!Content
        End If
        
        rs1.MoveNext
    
    Loop
    Set rs1 = Nothing
    
    
    nCount = 0
    
    sQry = "SELECT * FROM tb_reg" '��/ȣ�� ���� ��������
    Set rs1 = New ADODB.Recordset
    rs1.Open sQry, adoConn
    Do While Not (rs1.EOF)
        
        sDong = rs1!DRIVER_DEPT
        sHo = rs1!DRIVER_CLASS
        sID = Trim(Format(LeftH(rs1!DRIVER_DEPT, 3), "000") & Format(LeftH(rs1!DRIVER_CLASS, 4), "0000"))
        
        If (LenH(sID) = 7) Then
            sQry = "SELECT ID FROM tb_guestreg_admin WHERE ID = '" & sID & "' " '��/ȣ�� ���� ��������
            Set rs2 = New ADODB.Recordset
            rs2.Open sQry, adoConn
            If rs2.EOF Then
                sQry = "INSERT INTO tb_guestreg_admin (VENDOR, SITE, NAME, ID, PASSWORD, CARNO, TEL, DRIVER_DEPT, DRIVER_CLASS, MAXPARKTIME, MAXPARKCOUNT, NOWPARKCOUNT,MAXGUESTVISITCOUNT, USE_YN, REG_DATE) "
                sQry = sQry & " VALUES (0,0, '" & rs1!DRIVER_NAME & "', '" & sID & "', '0000', '" & rs1!CAR_NO & "', '" & rs1!DRIVER_PHONE & "', '" & sDong & "', '" & sHo & "', " & nParkTime & ", " & nParkCount & ", 0, '" & nVisitCount & "', 'Y', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "')"
                adoConn.Execute sQry

                Call DataLogger("[�����湮��û] �ű� ���̵� ���� : " & sID)
                
                nCount = nCount + 1
                
            End If
            Set rs2 = Nothing
        End If
        
        rs1.MoveNext
    Loop
    
    Call DataLogger("[�����湮��û] ID �ڵ� ���� ����")
        
    Call Clear_Field
    Call ListView_GuestRegCar_Draw
    Call ListView_GuestRegCar_SQL("SELECT * From tb_guestReg_admin ")
    
    Msg_Box.Label2.Caption = "[�����湮��û]"
    Msg_Box.Label1.Caption = "ID �ڵ� ���� (" & nCount & ")�� " & vbCrLf & vbCrLf & "�Ϸ��߽��ϴ�."
    Msg_Box.Show 1
    
    Exit Sub

Err_p:
    Call DataLogger("[�����湮��û] ID �ڵ� ���� ����:" & Err.Description)
End Sub




Private Sub txt_MaxParkDay_KeyPress(KeyAscii As Integer)
    '�������Է�
    If (txt_MaxParkDay = "0") Then
        txt_MaxParkDay = ""
    End If

    If (KeyAscii = 45) Then ' -
        txt_MaxParkDay = ""
    ElseIf (KeyAscii = vbKeyBack Or (KeyAscii >= vbKey0 And KeyAscii <= vbKey9)) Then '�齺���̽�, ����
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txt_MaxParkCount_KeyPress(KeyAscii As Integer)
    '�������Է�
    If (txt_MaxParkCount = "0") Then
        txt_MaxParkCount = ""
    End If

    If (KeyAscii = 45) Then ' -
        txt_MaxParkCount = ""
    ElseIf (KeyAscii = vbKeyBack Or (KeyAscii >= vbKey0 And KeyAscii <= vbKey9)) Then '�齺���̽�, ����
    Else
        KeyAscii = 0
    End If
End Sub


Private Sub txt_MaxParkTime_KeyPress(KeyAscii As Integer)
    '�������Է�
    If (txt_MaxParkTime = "0") Then
        txt_MaxParkTime = ""
    End If

    If (KeyAscii = 45) Then ' -
        txt_MaxParkTime = ""
    ElseIf (KeyAscii = vbKeyBack Or (KeyAscii >= vbKey0 And KeyAscii <= vbKey9)) Then '�齺���̽�, ����
    Else
        KeyAscii = 0
    End If
End Sub


Private Sub txt_NowParkCount_KeyPress(KeyAscii As Integer)
    '�������Է�
    If (txt_NowParkCount = "0") Then
        txt_NowParkCount = ""
    End If

    If (KeyAscii = 45) Then ' -
        txt_NowParkCount = ""
    ElseIf (KeyAscii = vbKeyBack Or (KeyAscii >= vbKey0 And KeyAscii <= vbKey9)) Then '�齺���̽�, ����
    Else
        KeyAscii = 0
    End If
End Sub


Private Sub txt_MaxGuestVisitCount_KeyPress(KeyAscii As Integer)
    '�������Է�
    If (txt_MaxGuestVisitCount = "0") Then
        txt_MaxGuestVisitCount = ""
    End If

    If (KeyAscii = 45) Then ' -
        txt_MaxGuestVisitCount = ""
    ElseIf (KeyAscii = vbKeyBack Or (KeyAscii >= vbKey0 And KeyAscii <= vbKey9)) Then '�齺���̽�, ����
    Else
        KeyAscii = 0
    End If
End Sub


