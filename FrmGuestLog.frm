VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMctl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmGuestLog 
   BackColor       =   &H00404040&
   BorderStyle     =   1  '���� ����
   Caption         =   "ParkingManager��"
   ClientHeight    =   11715
   ClientLeft      =   5160
   ClientTop       =   1725
   ClientWidth     =   17190
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   11715
   ScaleWidth      =   17190
   Begin VB.ComboBox cmb_Sort 
      DataField       =   "����"
      DataSource      =   "Data1(9)"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   19140
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   11
      Top             =   3810
      Width           =   1950
   End
   Begin VB.TextBox txt_CarNo 
      BeginProperty Font 
         Name            =   "�������"
         Size            =   20.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      IMEMode         =   10  '�ѱ� 
      Left            =   11460
      MaxLength       =   10
      TabIndex        =   10
      Top             =   4020
      Width           =   2805
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   " �ߺ� �湮 ���� �˻� "
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
      Height          =   1935
      Left            =   17340
      TabIndex        =   4
      Top             =   1395
      Width           =   6780
      Begin MSComCtl2.DTPicker DTPicker5 
         Height          =   345
         Left            =   405
         TabIndex        =   5
         ToolTipText     =   "������¥"
         Top             =   405
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   12648447
         CalendarForeColor=   12582912
         CalendarTitleBackColor=   8421504
         CalendarTitleForeColor=   12632256
         CalendarTrailingForeColor=   8421504
         Format          =   75104256
         CurrentDate     =   36927
      End
      Begin MSComCtl2.DTPicker DTPicker6 
         Height          =   345
         Left            =   3105
         TabIndex        =   6
         ToolTipText     =   "������¥"
         Top             =   405
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarBackColor=   12648447
         CalendarForeColor=   12582912
         CalendarTitleBackColor=   8421504
         CalendarTitleForeColor=   12632256
         CalendarTrailingForeColor=   8421504
         Format          =   75104256
         CurrentDate     =   36927
      End
      Begin Threed.SSCommand SSCommand3 
         Height          =   480
         Left            =   5430
         TabIndex        =   7
         Top             =   1290
         Width           =   1185
         _Version        =   65536
         _ExtentX        =   2090
         _ExtentY        =   847
         _StockProps     =   78
         Caption         =   "�� ��"
         ForeColor       =   16777215
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
         Picture         =   "FrmGuestLog.frx":0000
      End
      Begin VB.Label lbl_option 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "����"
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
         Height          =   285
         Index           =   15
         Left            =   5145
         TabIndex        =   9
         Top             =   435
         Width           =   450
      End
      Begin VB.Label lbl_option 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "����"
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
         Height          =   285
         Index           =   14
         Left            =   2445
         TabIndex        =   8
         Top             =   435
         Width           =   450
      End
   End
   Begin VB.ComboBox cmb_GuestDong 
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   11460
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   3
      Top             =   2925
      Width           =   1320
   End
   Begin VB.ComboBox cmb_GuestHo 
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   13590
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   2
      Top             =   2925
      Width           =   1320
   End
   Begin VB.ComboBox cmb_OrderBy 
      DataField       =   "����"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      ItemData        =   "FrmGuestLog.frx":0351
      Left            =   11460
      List            =   "FrmGuestLog.frx":0353
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   1
      Top             =   3495
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.TextBox txt_Count 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1042
         SubFormatType   =   0
      EndProperty
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
      Left            =   11430
      TabIndex        =   0
      ToolTipText     =   "1 �̻� ���� �Է�. ���� ���� �˻��� ó���ð��� �����ɼ��ֽ��ϴ�."
      Top             =   2175
      Width           =   525
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   6180
      Left            =   150
      TabIndex        =   12
      Top             =   5370
      Width           =   16890
      _ExtentX        =   29792
      _ExtentY        =   10901
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
   Begin Threed.SSCommand SSCommand1 
      Height          =   570
      Left            =   14550
      TabIndex        =   13
      Top             =   75
      Width           =   1185
      _Version        =   65536
      _ExtentX        =   2090
      _ExtentY        =   1005
      _StockProps     =   78
      Caption         =   "����"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "FrmGuestLog.frx":0355
   End
   Begin Threed.SSCommand SSCommand2 
      Cancel          =   -1  'True
      Height          =   570
      Left            =   15810
      TabIndex        =   14
      Top             =   75
      Width           =   1185
      _Version        =   65536
      _ExtentX        =   2090
      _ExtentY        =   1005
      _StockProps     =   78
      Caption         =   "�ݱ�"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "FrmGuestLog.frx":06A6
   End
   Begin Threed.SSPanel PnlOut 
      Height          =   390
      Index           =   7
      Left            =   11520
      TabIndex        =   15
      Top             =   4935
      Width           =   2520
      _Version        =   65536
      _ExtentX        =   4445
      _ExtentY        =   688
      _StockProps     =   15
      Caption         =   " �˻� �Ǽ�"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      Alignment       =   1
      Begin VB.Label LblRecordCount 
         Alignment       =   2  '��� ����
         BackColor       =   &H00000000&
         Caption         =   "000"
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
         Left            =   1170
         TabIndex        =   16
         Top             =   60
         Width           =   1275
      End
   End
   Begin Threed.SSCommand Command1 
      Height          =   585
      Left            =   15420
      TabIndex        =   17
      Top             =   4020
      Width           =   1620
      _Version        =   65536
      _ExtentX        =   2857
      _ExtentY        =   1032
      _StockProps     =   78
      Caption         =   "�� ��"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�������"
         Size            =   14.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "FrmGuestLog.frx":09F7
   End
   Begin Threed.SSPanel PnlOut 
      Height          =   390
      Index           =   0
      Left            =   14520
      TabIndex        =   18
      Top             =   4935
      Width           =   2520
      _Version        =   65536
      _ExtentX        =   4445
      _ExtentY        =   688
      _StockProps     =   15
      Caption         =   " ��  �����ð�"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelOuter      =   0
      RoundedCorners  =   0   'False
      Outline         =   -1  'True
      Alignment       =   1
      Begin VB.Label LblTotalParkTime 
         Alignment       =   2  '��� ����
         BackColor       =   &H00000000&
         Caption         =   "000"
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
         Left            =   1170
         TabIndex        =   19
         Top             =   60
         Width           =   1275
      End
   End
   Begin Threed.SSCommand SSCommand4 
      Height          =   570
      Left            =   12780
      TabIndex        =   20
      ToolTipText     =   "���� ������ȭ������ ��ȯ�մϴ�"
      Top             =   75
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   1005
      _StockProps     =   78
      Caption         =   "���� ��������ȸ"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "FrmGuestLog.frx":0D48
   End
   Begin Threed.SSCommand SSCommand5 
      Height          =   570
      Left            =   9660
      TabIndex        =   21
      ToolTipText     =   "���� ������ ��ȸ�� ��ü �ý��ۿ� ������ ��ĥ�� �ֽ��ϴ�."
      Top             =   75
      Width           =   1485
      _Version        =   65536
      _ExtentX        =   2619
      _ExtentY        =   1005
      _StockProps     =   78
      Caption         =   "���� �湮��ȸ"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "FrmGuestLog.frx":1099
   End
   Begin Threed.SSCommand SSCommand6 
      Height          =   570
      Left            =   11220
      TabIndex        =   22
      Top             =   75
      Width           =   1485
      _Version        =   65536
      _ExtentX        =   2619
      _ExtentY        =   1005
      _StockProps     =   78
      Caption         =   "���� �湮��ȸ"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "FrmGuestLog.frx":13EA
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   345
      Left            =   11415
      TabIndex        =   35
      Top             =   1605
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�������"
         Size            =   9
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   12648447
      CalendarForeColor=   12582912
      CalendarTitleBackColor=   8421504
      CalendarTitleForeColor=   12632256
      CalendarTrailingForeColor=   8421504
      Format          =   139526144
      CurrentDate     =   36927
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   345
      Left            =   14250
      TabIndex        =   36
      Top             =   1605
      Width           =   1950
      _ExtentX        =   3440
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�������"
         Size            =   9
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   12648447
      CalendarForeColor=   12582912
      CalendarTitleBackColor=   8421504
      CalendarTitleForeColor=   12632256
      CalendarTrailingForeColor=   8421504
      Format          =   139526144
      CurrentDate     =   36927
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9060
      Top             =   165
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '����
      DataField       =   "imgpath1"
      DataSource      =   "Adodc1"
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   255
      TabIndex        =   34
      Top             =   13410
      Width           =   14715
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "���ļ���"
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
      Height          =   285
      Index           =   4
      Left            =   18000
      TabIndex        =   33
      Top             =   3810
      Width           =   900
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "��ȸ�Ⱓ"
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
      Height          =   285
      Index           =   5
      Left            =   10380
      TabIndex        =   32
      Top             =   1635
      Width           =   900
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "����"
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
      Height          =   285
      Index           =   7
      Left            =   13455
      TabIndex        =   31
      Top             =   1635
      Width           =   450
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "����"
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
      Height          =   285
      Index           =   9
      Left            =   16305
      TabIndex        =   30
      Top             =   1635
      Width           =   450
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "������ȣ"
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
      Height          =   315
      Index           =   10
      Left            =   10320
      TabIndex        =   29
      Top             =   4140
      Width           =   1080
   End
   Begin VB.Label lbl_APS 
      BackStyle       =   0  '����
      Caption         =   " �湮�� �߱޳��� ��ȸ"
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
      Left            =   180
      TabIndex        =   28
      Top             =   210
      Width           =   4185
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   0
      X1              =   150
      X2              =   17010
      Y1              =   690
      Y2              =   690
   End
   Begin VB.Image Image3 
      Height          =   3750
      Index           =   0
      Left            =   165
      Picture         =   "FrmGuestLog.frx":173B
      Stretch         =   -1  'True
      Top             =   1455
      Width           =   4920
   End
   Begin VB.Image Image3 
      Height          =   3750
      Index           =   1
      Left            =   5100
      Picture         =   "FrmGuestLog.frx":EB08
      Stretch         =   -1  'True
      Top             =   1455
      Width           =   4920
   End
   Begin VB.Label lbl_GuestDong 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "ȸ���"
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
      Height          =   285
      Left            =   10470
      TabIndex        =   27
      Top             =   2955
      Width           =   675
   End
   Begin VB.Label lbl_GuestHo 
      Alignment       =   1  '������ ����
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�μ�"
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
      Height          =   285
      Left            =   12960
      TabIndex        =   26
      Top             =   2955
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "���ļ���"
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
      Height          =   285
      Left            =   10380
      TabIndex        =   25
      Top             =   3495
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�湮 ���� �˻�"
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
      Height          =   285
      Index           =   17
      Left            =   13050
      TabIndex        =   24
      Top             =   2205
      Width           =   1470
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "ȸ �̻�"
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
      Height          =   285
      Index           =   16
      Left            =   12090
      TabIndex        =   23
      Top             =   2205
      Width           =   735
   End
End
Attribute VB_Name = "FrmGuestLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim excel_sql_str As String
Dim Last_Search_Index As Long

Public Guest_Search As String
Public Guest_Old As Boolean
Public Gueset_Old_EndDate As String
Public OutGate_YN As Boolean '�ⱸ �ִ��� Ȯ��


Private Sub DetailMultiSearch()
    Dim i As Integer
    Dim Cnt As Integer
    Dim Current_Date As String
    Dim TmpPath As String
    Dim Tmp_File As String
    Dim InsSQL As String
    Dim Now_Flag As Boolean
    Dim sql_str As String
    Dim Sort_Order As String

On Error Resume Next
    Me.MousePointer = 11
        
        Glo_SQL_SEARCH = ""
        
        If (Guest_Old = True) Then
            If (DTPicker2 > Gueset_Old_EndDate) Then
                DTPicker2 = Gueset_Old_EndDate
            End If
        End If
'            sql_str = "SELECT * FROM tb_guest_log_backup WHERE (REG_DATE >= '" & Format(DTPicker1, "yyyy-mm-dd") & " " & "00:00:00.000" & "') AND (REG_DATE <= '" & Format(DTPicker2, "yyyy-mm-dd") & " " & "23:59:59.999" & "')"
'        Else
'            sql_str = "SELECT * FROM tb_guest_log WHERE (REG_DATE >= '" & Format(DTPicker1, "yyyy-mm-dd") & " " & "00:00:00.000" & "') AND (REG_DATE <= '" & Format(DTPicker2, "yyyy-mm-dd") & " " & "23:59:59.999" & "')"
'        End If

        If (OutGate_YN = True) Then '�ⱸ�� �������
            sql_str = "SELECT CAR_NO, IN_COUNT, PARKTIME From (SELECT tb_guest_log.`CAR_NO` AS CAR_NO, count(*) AS IN_COUNT, SUM(PARK_TIME) AS PARKTIME From tb_guest_log Where tb_guest_log.DT_IN >= '" & Format(DTPicker1, "yyyy-mm-dd") & " 00:00:00.000' AND tb_guest_log.DT_OUT <= '" & Format(DTPicker2, "yyyy-mm-dd") & " 23:59:59.999' Group By tb_guest_log.`CAR_NO`) AS PARKONE Where IN_COUNT >= " & Val(txt_Count.text) & ""
        Else
            sql_str = "SELECT CAR_NO, IN_COUNT, PARKTIME From (SELECT tb_guest_log.`CAR_NO` AS CAR_NO, count(*) AS IN_COUNT, SUM(PARK_TIME) AS PARKTIME From tb_guest_log Where tb_guest_log.DT_IN >= '" & Format(DTPicker1, "yyyy-mm-dd") & " 00:00:00.000' AND tb_guest_log.DT_IN <= '" & Format(DTPicker2, "yyyy-mm-dd") & " 23:59:59.999' Group By tb_guest_log.`CAR_NO`) AS PARKONE Where IN_COUNT >= " & Val(txt_Count.text) & ""
        End If
        
        '������ȣ �˻�
        If (txt_CarNo.text = "") Then
        Else
            sql_str = sql_str & " AND (CAR_NO = '" & txt_CarNo.text & "')"
        End If

        'Sort_Order = " ORDER BY REG_DATE"
        'sql_str = sql_str & Sort_Order
        Glo_SQL_SEARCH = sql_str
        Call ListView_Draw(sql_str)
        
        Me.MousePointer = 0
End Sub

Private Sub SingleSearch()
    Dim i As Integer
    Dim Cnt As Integer
    Dim Current_Date As String
    Dim TmpPath As String
    Dim Tmp_File As String
    Dim InsSQL As String
    Dim Now_Flag As Boolean
    Dim sql_str As String
    Dim Sort_Order As String

On Error Resume Next
    Me.MousePointer = 11
        
        Glo_SQL_SEARCH = ""
        
        If (Guest_Old = True) Then
            If (DTPicker2 > Gueset_Old_EndDate) Then
                DTPicker2 = Gueset_Old_EndDate
            End If
            sql_str = "SELECT * FROM tb_guest_log_backup WHERE (REG_DATE >= '" & Format(DTPicker1, "yyyy-mm-dd") & " " & "00:00:00" & "') AND (REG_DATE <= '" & Format(DTPicker2, "yyyy-mm-dd") & " " & "23:59:59" & "')"
        Else
            sql_str = "SELECT * FROM tb_guest_log WHERE (REG_DATE >= '" & Format(DTPicker1, "yyyy-mm-dd") & " " & "00:00:00" & "') AND (REG_DATE <= '" & Format(DTPicker2, "yyyy-mm-dd") & " " & "23:59:59" & "')"
    '        Debug.Print sql_str
        End If
        
        '������ȣ �˻�
        If (txt_CarNo.text = "") Then
        Else
            sql_str = sql_str & " AND (CAR_NO Like '%" & txt_CarNo.text & "%')"
        End If
        
'        If (cmb_GuestDong.text <> "��ü") Then
'            sql_str = sql_str & " AND (Dong = '" & cmb_GuestDong.text & "') "
'        End If
'        If (cmb_GuestHo.text <> "��ü") Then
'            sql_str = sql_str & " AND (Ho = '" & cmb_GuestHo.text & "') "
'        End If

'''        Select Case cmb_Sort.ListIndex
'''            Case 0
'''                Sort_Order = "REG_DATE ASC"
'''            Case 1
'''                Sort_Order = "REG_DATE DESC"
'''        End Select
        Sort_Order = " ORDER BY CAR_NO, REG_DATE"
        
        sql_str = sql_str & Sort_Order
        Glo_SQL_SEARCH = sql_str
        Call ListView_Draw(sql_str)
        
        Me.MousePointer = 0
        
End Sub

Private Sub MultiSearch()
    Dim i As Integer
    Dim Cnt As Integer
    Dim Current_Date As String
    Dim TmpPath As String
    Dim Tmp_File As String
    Dim InsSQL As String
    Dim Now_Flag As Boolean
    Dim sql_str As String
    Dim Sort_Order As String

On Error Resume Next
    Me.MousePointer = 11
        Glo_SQL_SEARCH = ""
        
        '���� ����
        If (Guest_Old = True) Then '���Ź湮��ȸ
            If (DTPicker2 > Gueset_Old_EndDate) Then
                DTPicker2 = Gueset_Old_EndDate
            End If
        ElseIf (Guest_Old = False) Then '����湮��ȸ
        End If
        
        sql_str = "SELECT CAR_NO, count(CAR_NO) as IN_COUNT From tb_guest_log Where GUBUN = '�Ա�' AND (REG_DATE >= '" & Format(DTPicker1, "yyyy-mm-dd") & " " & "00:00:00" & "') AND (REG_DATE <= '" & Format(DTPicker2, "yyyy-mm-dd") & " " & "23:59:59" & "') "
        
        '������ȣ �˻�
        If (txt_CarNo.text = "") Then
        Else
            sql_str = sql_str & " AND (CAR_NO Like '%" & txt_CarNo.text & "%')"
        End If
        
        If (cmb_GuestDong.text <> "��ü") Then
            sql_str = sql_str & " AND (Dong = '" & cmb_GuestDong.text & "') "
        End If
        If (cmb_GuestHo.text <> "��ü") Then
            sql_str = sql_str & " AND (Ho = '" & cmb_GuestHo.text & "') "
        End If
        sql_str = sql_str & " GROUP BY CAR_NO"
        
        
'''        Sort_Order = ""
'''        Select Case cmb_OrderBy.text
'''            Case "�湮Ƚ��"
'''                Sort_Order = " ORDER BY IN_COUNT DESC"
'''            Case "�����ð�"
'''                Sort_Order = " ORDER BY PARKTIME DESC"
'''        End Select
'''        sql_str = sql_str & Sort_Order
         
        Glo_SQL_SEARCH = sql_str
        Call ListView_Draw_MultiSearch(sql_str)
        
        Me.MousePointer = 0

End Sub

'�˻� ����
Private Sub Command1_Click()
    Dim i As Integer
    Dim Cnt As Integer
    Dim Current_Date As String
    Dim TmpPath As String
    Dim Tmp_File As String
    Dim InsSQL As String
    Dim Now_Flag As Boolean
    Dim sql_str As String
    Dim Sort_Order As String

On Error Resume Next
    
    MousePointer = 11
    
    If IsNumeric(txt_Count.text) Then
        If (txt_Count.text <= 0) Then
            MsgBox " �ùٸ� ���ڸ� �Է��ϼ���...!! "
            Me.MousePointer = 0
            Exit Sub
        End If
    Else
        MsgBox " ���ڸ� �Է��ϼ���...!! "
        Me.MousePointer = 0
        Exit Sub
    End If
    
'''
'''    If (Val(txt_Count.text) > 1) Then
'''        Guest_Search = "�ߺ��˻�"
'''        Call MultiSearch
'''    Else
'''        Guest_Search = "�Ϲݰ˻�"
'''        Call SingleSearch
'''    End If

    Guest_Search = "�ߺ��˻�"
    Call MultiSearch
    
'
'    If (Guest_Search = "�ߺ��˻�") Then
'        Call MultiSearch
'
'    Else
'        Call SingleSearch
'
'    End If

    'SSPanel3(1).Caption = ""
    Me.MousePointer = 0

End Sub


Private Sub Form_Load()
    Dim Record_Source As String
    Dim i As Integer
    
'On Error GoTo err_P
    
    Left = (Screen.width - width) / 2   ' ���� ���η� �߾ӿ� �����ϴ�.
    Top = (Screen.height - height) / 2   ' ���� ���η� �߾ӿ� �����ϴ�.
    
    With Me.cmb_Sort
        .AddItem "�ð���"
        .AddItem "�ð�����"
        .text = Me.cmb_Sort.List(0)
    End With
    Me.cmb_Sort = Me.cmb_Sort.List(0)
    
    DTPicker1.value = Now
    DTPicker2.value = Now
    
    DTPicker5.value = Now
    DTPicker6.value = Now
    
    Image3(0).Picture = LoadPicture(App.Path & "\NoCar.jpg")
    Image3(1).Picture = LoadPicture(App.Path & "\NoCar.jpg")
    
    '���系����ȸ��ư
    SSCommand5.ForeColor = &HFFFFFF
    SSCommand6.ForeColor = &H808080
    
    txt_Count.text = 10
    cmb_OrderBy.AddItem "�湮Ƚ��"
    cmb_OrderBy.AddItem "�����ð�"
    cmb_OrderBy.text = "�湮Ƚ��"
    
    
    If (Glo_User_Type = "����1/����2") Then
        lbl_GuestDong = "�Ҽ�"
        lbl_GuestHo = "����"
    Else
        lbl_GuestDong = "��"
        lbl_GuestHo = "ȣ"
    End If
    Call SetGuestDong
    Call SetGuestHo
    OutGate_YN = FindOutGate
    
    
    '�˻� �ڷ� ���Բ�..
    'Glo_SQL_SEARCH = "SELECT * FROM tb_guest_log WHERE (REG_DATE >= '" & Format(DTPicker1, "yyyy-mm-dd") & " 00:00:00 000') AND (REG_DATE <= '" & Format(DTPicker2, "yyyy-mm-dd") & " 00:00:00 000')"
    Glo_SQL_SEARCH = "SELECT * FROM tb_guest_log WHERE idx=-1"
    
    Guest_Search = "�Ϲݰ˻�"
    
    'Call ListView_Draw
    
    
    Call Set_VisitQueryTerm_ToolTip
    

Exit Sub
    
Err_p:
    MsgBox "������ ���̽� �������" & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & "��Ʈ�� �����ڿ��� ���� �ٶ��ϴ�." & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & "������ ���̽� ���������� �ڷ�˻� ����� �����Ҽ� �����ϴ�.", vbCritical
End Sub

Private Function FindOutGate() As Boolean
    If ((LANE1_YN = "Y" And LANE1_Inout = "�ⱸ") Or (LANE2_YN = "Y" And LANE2_Inout = "�ⱸ") Or (LANE3_YN = "Y" And LANE3_Inout = "�ⱸ") Or (LANE4_YN = "Y" And LANE4_Inout = "�ⱸ") Or (LANE5_YN = "Y" And LANE5_Inout = "�ⱸ") Or (LANE6_YN = "Y" And LANE6_Inout = "�ⱸ")) Then
        FindOutGate = True
    Else
        FindOutGate = False
    End If
End Function

Private Sub Set_VisitQueryTerm_ToolTip()

    Dim rs As Recordset
    Dim qry As String
    Dim bQryResult As Boolean

On Error GoTo Err_p

    Set rs = New ADODB.Recordset
    bQryResult = DataBaseQuery(rs, adoConn, "Select min(reg_date) as MinDate, max(reg_date) as MaxDate from tb_guest_log", False)
    If (bQryResult = False) Then
        Call DataLogger("[FrmGuestLog VisitQueryTerm]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
        Set rs = Nothing
        Exit Sub
    End If
    If (Not rs.EOF) Then
        SSCommand6.ToolTipText = "�˻��Ⱓ:" & Left(rs!MinDate, 10) & "~" & Left(rs!MaxDate, 10) '����湮����ȸ �Ⱓ
    End If
    Set rs = Nothing
    
    
    
    Set rs = New ADODB.Recordset
    bQryResult = DataBaseQuery(rs, adoConn, "Select count(*), max(reg_date) as MaxDate from tb_guest_log_backup", False)
    If (bQryResult = False) Then
        Call DataLogger("[FrmGuestLog VisitQueryTerm]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
        Set rs = Nothing
        Exit Sub
    End If
    If (Not rs.EOF) Then
        SSCommand5.Enabled = False
        SSCommand5.Visible = False
        SSCommand5.ToolTipText = "" '���Ź湮����ȸ �Ⱓ
    Else
        SSCommand5.Enabled = True
        SSCommand5.Visible = True
        SSCommand5.ToolTipText = "�˻��Ⱓ:" & "����~" & Left(rs!MaxDate, 10) '���Ź湮����ȸ �Ⱓ
    End If
    Set rs = Nothing
    

    Exit Sub
    
Err_p:
    Set rs = Nothing
End Sub

Private Sub SetGuestDong()
    Dim bQryResult As Boolean
    Dim rs As Recordset
    Dim qry As String
    
    cmb_GuestDong.Clear
    cmb_GuestDong.AddItem "��ü"
    
    qry = "SELECT DONG From tb_guest_log Group By DONG"
    'QRY = "SELECT DONG From tb_guest_log Where tb_guest_log.DT_IN >= '" & Format(DTPicker5, "yyyy-mm-dd") & " 00:00:00' Group By DONG"
    Set rs = New ADODB.Recordset
    bQryResult = DataBaseQuery(rs, adoConn, qry, False)

    If Not rs.EOF Then
        Do While Not (rs.EOF)
            cmb_GuestDong.AddItem "" & rs!Dong
            rs.MoveNext
        Loop
    End If
    Set rs = Nothing
    
    cmb_GuestDong.ListIndex = 0
    'cmb_GuestDong.Refresh
End Sub

Private Sub SetGuestHo()
    Dim bQryResult As Boolean
    Dim rs As Recordset
    Dim qry As String
    
    
    cmb_GuestHo.Clear
    cmb_GuestHo.AddItem "��ü"
    
    qry = "SELECT Ho From tb_guest_log Group By Ho"
    Set rs = New ADODB.Recordset
    bQryResult = DataBaseQuery(rs, adoConn, qry, False)
    
    If Not rs.EOF Then
        Do While Not (rs.EOF)
            cmb_GuestHo.AddItem "" & rs!Ho
            rs.MoveNext
        Loop
    End If
    Set rs = Nothing
    
    cmb_GuestHo.ListIndex = 0
End Sub



Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
    Dim i As Integer
    With ListView1
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

Private Sub ListView1_ItemClick(ByVal Item As ComctlLib.ListItem)
    Dim Tmp_File As String
    Dim image_name As String
    Dim i As Integer
    Dim TestIP As String
    Dim ECHO As ICMP_ECHO_REPLY
    
    On Error Resume Next
    
    If (Guest_Search = "�Ϲݰ˻�") Then
    
        If (IsFile(ListView1.SelectedItem.SubItems(12)) = True) Then
            Image3(0).Picture = LoadPicture(ListView1.SelectedItem.SubItems(12))
        Else
            Image3(0).Picture = LoadPicture(App.Path & "\NoCar.jpg")
        End If
        
        If (IsFile(ListView1.SelectedItem.SubItems(13)) = True) Then
            Image3(1).Picture = LoadPicture(ListView1.SelectedItem.SubItems(13))
        Else
            Image3(1).Picture = LoadPicture(App.Path & "\NoCar.jpg")
        End If
        
    Else
        Guest_Search = "�Ϲݰ˻�"
        txt_CarNo = Trim(ListView1.SelectedItem.SubItems(1))
        '''Call DetailMultiSearch
        Call SingleSearch
    
    End If

    txt_CarNo = ""

End Sub


Private Sub SSCommand1_Click()
'''Dim myExcelFile As New ExcelFile
'''Dim tmpFileName As String
'''
'''tmpFileName = Format(Now, "YYYYMMDD_HHMMSS")
'''
'''If (Guest_Search = "�Ϲݰ˻�") Then
'''    tmpFileName = App.Path & "\Excel\" & tmpFileName & "_�湮�� �߱� �˻�����" & ".xls"
'''    'Call makeexcel(ListView1, tmpFileName, "�湮�� �߱� �˻�����")
'''    Call MakeCSV(ListView1, tmpFileName)
'''Else
'''    tmpFileName = App.Path & "\Excel\" & tmpFileName & "_�湮���� Ư�� �˻�����" & ".xls"
'''    'Call makeexcel(ListView1, tmpFileName, "�湮���� Ư�� �˻�����")
'''    Call MakeCSV(ListView1, tmpFileName)
'''End If
'''

    Dim i, j As Integer
    Dim myExcelFile As New ExcelFile
    Dim tmpFileName As String
    
On Error GoTo Err_p
    tmpFileName = Format(Now, "YYYYMMDD_HHMMSS")
    If (Guest_Search = "�Ϲݰ˻�") Then
        tmpFileName = App.Path & "\Excel\" & tmpFileName & "_�湮�� �߱� �˻�����" & ".xls"
    Else
        tmpFileName = App.Path & "\Excel\" & tmpFileName & "_�湮���� Ư�� �˻�����" & ".xls"
    End If
        
    CommonDialog1.CancelError = True
    CommonDialog1.InitDir = App.Path
    CommonDialog1.Filter = "��������(*.csv)|*.csv"
    CommonDialog1.fileName = tmpFileName
    CommonDialog1.ShowSave

    If (CommonDialog1.CancelError = True) Then
    
        tmpFileName = CommonDialog1.fileName
        tmpFileName = Mid(tmpFileName, 1, Len(tmpFileName) - 4)
        Call MakeCSV(ListView1, tmpFileName)
    End If
Exit Sub

Err_p:
     Select Case Err
    Case 32755 '  Dialog Cancelled
        'MsgBox "you cancelled the dialog box"
    Case Else
        'MsgBox "Unexpected error. Err " & Err & " : " & Error
    End Select
    

End Sub

'����
Private Sub SSCommand2_Click()
    'Unload Me
    Me.Hide
End Sub

Public Sub ListView_Draw(sQry As String)
Dim Column_to_size As Integer
Dim rs As Recordset
Dim itmX As ListItem
Dim INDEX_NO, col As Long
Dim totalParkTime As Long

Dim sLastGubun As String

Dim sInCarno As String
Dim sInTime As String
Dim sInGate As String
Dim sInImage As String
Dim sInRegDate As String
Dim sName As String
Dim sDong  As String
Dim sHo  As String
Dim sTel  As String
Dim sObject  As String
            
Dim sOutCarno As String
Dim sOutTime As String
Dim sOutGate As String
Dim sOutImage As String
Dim sOutRegDate As String
Dim nParkTM As Long

Dim guest As stGuest

    Call ListViewExtended(ListView1)
    ListView1.View = lvwReport
    ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , " No "
    ListView1.ColumnHeaders.Add , , " ������ȣ         "
    ListView1.ColumnHeaders.Add , , " �湮��       "
    
    If (Glo_User_Type = "����1/����2") Then
        ListView1.ColumnHeaders.Add , , " ��   ��1     "
        ListView1.ColumnHeaders.Add , , " ��   ��2     "
    Else
        ListView1.ColumnHeaders.Add , , " ��           "
        ListView1.ColumnHeaders.Add , , " ȣ           "
    End If
    
    ListView1.ColumnHeaders.Add , , " ����ó           "
    ListView1.ColumnHeaders.Add , , " �� ��        "
    ListView1.ColumnHeaders.Add , , " �����ð�(��)"
    
    ListView1.ColumnHeaders.Add , , " ��������Ʈ     "
    ListView1.ColumnHeaders.Add , , " ������¥                     "
    ListView1.ColumnHeaders.Add , , " ��������Ʈ     "
    ListView1.ColumnHeaders.Add , , " ������¥       "
    
    ListView1.ColumnHeaders.Add , , " �����̹������            "
    ListView1.ColumnHeaders.Add , , " �����̹������            "
    
    ListView1.ColumnHeaders.Add , , " �� ��                        "
    ListView1.ColumnHeaders.Add , , " �����           "
    
    
    For Column_to_size = 0 To ListView1.ColumnHeaders.Count - 1
         SendMessage ListView1.hwnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
    Next
    
    Set rs = New ADODB.Recordset
    rs.Open sQry, adoConn
    
'    rs.Open "Select * from tb_guest_log order by car_no, reg_date", adoConn
'''    LblRecordCount = rs.RecordCount

    INDEX_NO = 1
    sLastGubun = ""
    
    Do While Not (rs.EOF)
    
            If (rs!Gubun = "�Ա�") Then

                guest.InCarNo = rs!CAR_NO
                guest.GuestName = rs!name
                guest.Dong = rs!Dong
                guest.Ho = rs!Ho
                guest.Tel = rs!Tel
                guest.object = rs!object
                guest.InDate = rs!DT_IN
                'guest.InGateNo = rs!IN_GATE
                guest.InImagePath = rs!IN_IMAGE_PATH
                guest.RegDate = rs!REG_DATE

                guest.ParkTime = 0
                guest.OutDate = ""
                guest.OutGateNo = ""
                guest.OutImagePath = ""
                
                If (rs!IN_GATE = "0") Then
                    guest.InGateNo = LANE1_Name
                ElseIf (rs!IN_GATE = "1") Then
                    guest.InGateNo = LANE2_Name
                ElseIf (rs!IN_GATE = "2") Then
                    guest.InGateNo = LANE3_Name
                ElseIf (rs!IN_GATE = "3") Then
                    guest.InGateNo = LANE4_Name
                ElseIf (rs!IN_GATE = "4") Then
                    guest.InGateNo = LANE5_Name
                ElseIf (rs!IN_GATE = "5") Then
                    guest.InGateNo = LANE6_Name
                Else
                    guest.InGateNo = ""
                End If


            ElseIf (rs!Gubun = "�ⱸ") Then
                guest.OutCarNo = rs!CAR_NO
                guest.OutDate = rs!DT_OUT
                'guest.OutGateNo = rs!OUT_GATE
                guest.OutImagePath = rs!OUT_IMAGE_PATH
                guest.RegDate = rs!REG_DATE
                
                If (rs!OUT_GATE = "0") Then
                    guest.OutGateNo = LANE1_Name
                ElseIf (rs!OUT_GATE = "1") Then
                    guest.OutGateNo = LANE2_Name
                ElseIf (rs!OUT_GATE = "2") Then
                    guest.OutGateNo = LANE3_Name
                ElseIf (rs!OUT_GATE = "3") Then
                    guest.OutGateNo = LANE4_Name
                ElseIf (rs!OUT_GATE = "4") Then
                    guest.OutGateNo = LANE5_Name
                ElseIf (rs!OUT_GATE = "5") Then
                    guest.OutGateNo = LANE6_Name
                Else
                    guest.OutGateNo = ""
                End If

                If (guest.InCarNo = guest.OutCarNo) Then
                    guest.ParkTime = DateDiff("n", Left(guest.InDate, 19), Left(guest.OutDate, 19))
                    totalParkTime = totalParkTime + guest.ParkTime
                End If


                Call Draw_Listview_Guest(guest, INDEX_NO)
                Call ClearGuestInfo(guest)
                
                
                INDEX_NO = INDEX_NO + 1
            End If
            rs.MoveNext
    Loop
    
    Set rs = Nothing
    
    If (Len(guest.InCarNo) > 0) Then
        Call Draw_Listview_Guest(guest, INDEX_NO)
        Call ClearGuestInfo(guest)
    End If
    
    LblRecordCount.Caption = INDEX_NO - 1
    LblTotalParkTime = totalParkTime

End Sub

Private Sub ClearGuestInfo(guest As stGuest)
    guest.InCarNo = ""
    guest.GuestName = ""
    guest.Dong = ""
    guest.Ho = ""
    guest.Tel = ""
    guest.object = ""
    guest.InGateNo = ""
    guest.InDate = ""
    guest.InImagePath = ""
    guest.RegDate = ""
    guest.ParkTime = ""
    
    guest.OutCarNo = ""
    guest.OutGateNo = ""
    guest.OutDate = ""
    guest.OutImagePath = ""
End Sub

Private Sub Draw_Listview_Guest(guest As stGuest, ByVal IndexNo As Integer)

    Dim itmX As ListItem
    Dim col As Integer
    
    Set itmX = ListView1.ListItems.Add(, , "" & IndexNo)

    col = 1
    itmX.SubItems(col) = "" & guest.InCarNo: col = col + 1
    itmX.SubItems(col) = "" & guest.GuestName: col = col + 1
    itmX.SubItems(col) = "" & guest.Dong: col = col + 1
    itmX.SubItems(col) = "" & guest.Ho: col = col + 1
    itmX.SubItems(col) = "" & guest.Tel: col = col + 1
    itmX.SubItems(col) = "" & guest.object: col = col + 1

    itmX.SubItems(col) = "" & guest.ParkTime: col = col + 1
    
    itmX.SubItems(col) = "" & guest.InGateNo: col = col + 1
    itmX.SubItems(col) = "" & guest.InDate: col = col + 1
    itmX.SubItems(col) = "" & guest.OutGateNo: col = col + 1
    itmX.SubItems(col) = "" & guest.OutDate: col = col + 1
    itmX.SubItems(col) = "" & guest.InImagePath: col = col + 1
    itmX.SubItems(col) = "" & guest.OutImagePath: col = col + 1
    
    itmX.SubItems(col) = "" & guest.object: col = col + 1
    itmX.SubItems(col) = "" & guest.RegDate: col = col + 1
    
End Sub

'����Ű �Է½� �� ����
'���Ӽ� keypreview = true ����
Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        Call Command1_Click
        KeyAscii = 0
        Sendkeys "{TAB}"
    End If

End Sub

'Custum Search
Private Sub SSCommand3_Click()
    Dim i As Integer
    Dim Cnt As Integer
    Dim Current_Date As String
    Dim TmpPath As String
    Dim Tmp_File As String
    Dim InsSQL As String
    Dim Now_Flag As Boolean
    Dim sql_str As String
    Dim Sort_Order As String
    
    Guest_Search = "�ߺ��˻�"
    
    Me.MousePointer = 11
    Glo_SQL_SEARCH = ""
    
    If IsNumeric(txt_Count.text) Then
        If (txt_Count.text = 0) Then
            MsgBox " �ùٸ� ���ڸ� �Է��ϼ���...!! "
            Me.MousePointer = 0
            Exit Sub
        Else
    
        End If
    Else
        MsgBox " ���ڸ� �Է��ϼ���...!! "
        Me.MousePointer = 0
        Exit Sub
    End If
    
    '���� ����
    'sql_str = "SELECT CAR_NO, IN_COUNT From (SELECT tb_guest_log.`CAR_NO` AS CAR_NO, count(*) AS IN_COUNT From tb_guest_log Where tb_guest_log.REG_DATE >= '" & Format(DTPicker5, "yyyy-mm-dd") & " 00:00:00' AND tb_guest_log.REG_DATE <= '" & Format(DTPicker6, "yyyy-mm-dd") & " 23:59:59' Group By tb_guest_log.`CAR_NO`) AS PARKONE Where IN_COUNT >= " & val(txt_Count.Text) & ""
    sql_str = "SELECT CAR_NO, IN_COUNT, PARKTIME From (SELECT tb_guest_log.`CAR_NO` AS CAR_NO, count(*) AS IN_COUNT, SUM(PARK_TIME) AS PARKTIME From tb_guest_log Where tb_guest_log.DT_IN >= '" & Format(DTPicker5, "yyyy-mm-dd") & " 00:00:00' AND tb_guest_log.DT_OUT <= '" & Format(DTPicker6, "yyyy-mm-dd") & " 23:59:59' Group By tb_guest_log.`CAR_NO`) AS PARKONE Where IN_COUNT >= " & Val(txt_Count.text) & ""
    'Debug.Print sql_str
    
    Sort_Order = ""
    Select Case cmb_OrderBy.text
        Case "�湮Ƚ��"
            Sort_Order = " ORDER BY IN_COUNT DESC"
        Case "�����ð�"
            Sort_Order = " ORDER BY PARKTIME DESC"
    End Select
    sql_str = sql_str & Sort_Order
     
    Glo_SQL_SEARCH = sql_str
    Call ListView_Draw_MultiSearch(sql_str)
    Me.MousePointer = 0

On Error Resume Next

End Sub

Public Sub ListView_Draw_MultiSearch(sQry As String)
    Dim Column_to_size As Integer
    Dim rs As Recordset
    Dim qry As String
    Dim itmX As ListItem
    Dim INDEX_NO As Long
    
    Call ListViewExtended(ListView1)
    ListView1.View = lvwReport
    ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , " No "
    ListView1.ColumnHeaders.Add , , " ������ȣ         "
    ListView1.ColumnHeaders.Add , , " �����Ǽ�(ȸ) "
    'ListView1.ColumnHeaders.Add , , " �����ð�(��)       "
    
    For Column_to_size = 0 To ListView1.ColumnHeaders.Count - 1
         SendMessage ListView1.hwnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
    Next
    
    Set rs = New ADODB.Recordset
'    rs.Open Glo_SQL_SEARCH, adoConn
    rs.Open sQry, adoConn
'    LblRecordCount = rs.RecordCount
'    If (rs.RecordCount > 0) Then
'        LblTotalParkTime = rs!PARKTIME
'    Else
'        LblTotalParkTime = 0
'    End If
    
    INDEX_NO = 1
    LblRecordCount = 0
    LblTotalParkTime = 0
    
    Do While Not (rs.EOF)
        If (rs!IN_COUNT >= Val(txt_Count.text)) Then
        
            Set itmX = ListView1.ListItems.Add(, , "" & INDEX_NO)
            itmX.SubItems(1) = "" & rs!CAR_NO
            itmX.SubItems(2) = "" & rs!IN_COUNT
            'itmX.SubItems(3) = "" & rs!PARKTIME
            
    '        ListView1.Refresh
            
            LblRecordCount = INDEX_NO
            LblRecordCount.Refresh
            
            'LblTotalParkTime = LblTotalParkTime + rs!ParkTime
            'LblTotalParkTime.Refresh
            
            INDEX_NO = INDEX_NO + 1
        End If
        
        rs.MoveNext
        

'''        If ((INDEX_NO Mod 10000) = 0) Then
'''            MBox.Label2.Caption = "�湮�߱� ���� ��ȸ"
'''            MBox.Label3.Caption = "�˻� �����Ͱ� ���� �ð��� ������ �� �ֽ��ϴ�."
'''            MBox.Label3.FontSize = 20
'''            MBox.Label1.Caption = "�˻� �����Ͱ� ���� �ð��� ������ �� �ֽ��ϴ�. ����Ͻðڽ��ϱ�?"
'''            MBox.Show 1
'''
'''            If (Glo_MsgRet = False) Then
'''                INDEX_NO = 0
'''                Set rs = Nothing
'''                Exit Sub
'''            End If
'''        End If
    Loop
'Debug.Print ("������� : " & Format(Now, "yyyy-mm-dd hh:nn:ss"))
    INDEX_NO = 0
    Set rs = Nothing

End Sub


Private Sub SSCommand4_Click()
    'Unload Me
    'FrmInOut.Show 1
    
    Me.Hide
    FrmInOut.Show 0
    Call DataLogger("[GUEST Button]    " & "���������� ȭ�� ����")
End Sub

'���� �湮�� ���� ��ȸ
Private Sub SSCommand5_Click()
    Dim bQryResult As Boolean
    Dim rs As Recordset
    Dim qry As String
    
    Guest_Old = True
    
    qry = "SELECT max(reg_date) as MaxDate, min(reg_date) as MinDate FROM tb_guest_log_backup order by reg_date DESC LIMIT 1"
    Set rs = New ADODB.Recordset
    bQryResult = DataBaseQuery(rs, adoConn, qry, False)

    If Not rs.EOF Then
        DTPicker1 = Format(Left(rs!MinDate, 19), "yyyy-mm-dd AMPM hh:mm:ss") '2019-01-01 ���� 09:30:00
    Else
        DTPicker1 = Now
    End If

    DTPicker2 = Format(Left(rs!MaxDate, 19), "yyyy-mm-dd AMPM hh:mm:ss")
    
    Gueset_Old_EndDate = DTPicker2
    Set rs = Nothing
    
    '���系����ȸ��ư
    SSCommand5.ForeColor = &H808080
    SSCommand6.ForeColor = &HFFFFFF
    
End Sub

Private Sub SSCommand6_Click()
    
    SSCommand5.ForeColor = &HFFFFFF
    SSCommand6.ForeColor = &H808080
    Guest_Old = False
    DTPicker1 = Now
    DTPicker2 = Now
End Sub

Private Sub txt_Count_Change()
    
    cmb_OrderBy.Clear
    
    If (Val(txt_Count.text) >= 2) Then

        cmb_OrderBy.AddItem "�湮Ƚ��"
        cmb_OrderBy.AddItem "�����ð�"
        cmb_OrderBy.text = "�湮Ƚ��"
    
    Else
        cmb_OrderBy.AddItem "�����ð�"
        cmb_OrderBy.text = "�����ð�"
    End If
    
    If (Val(txt_Count.text) < 1) Then
        txt_Count.text = "0"
    End If
    
End Sub


