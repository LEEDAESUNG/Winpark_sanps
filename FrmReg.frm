VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMctl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmReg 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  '���� ����
   Caption         =   "ParkingManager��"
   ClientHeight    =   12885
   ClientLeft      =   5730
   ClientTop       =   2100
   ClientWidth     =   15375
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   Picture         =   "FrmReg.frx":0000
   ScaleHeight     =   12885
   ScaleWidth      =   15375
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chk_App 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   17430
      TabIndex        =   83
      ToolTipText     =   "����Ͼ� üũ������ ���, ��й�ȣ �ʱ�ȭ�ǹǷ� �ݵ�� ���ѳ��� �����ؾ��մϴ�."
      Top             =   8955
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.CommandButton cmd_PWInit 
      Caption         =   "��й�ȣ�ʱ�ȭ"
      Height          =   495
      Left            =   17730
      TabIndex        =   82
      ToolTipText     =   "����Ͼ� ���� ��й�ȣ �ʱ�ȭ ��, �ݵ�� ���ѳ��� �����ؾ� �մϴ�(�ʱ� ��й�ȣ 12345678)"
      Top             =   8745
      Visible         =   0   'False
      Width           =   975
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
      Left            =   15975
      TabIndex        =   70
      Top             =   3930
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1215
   End
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
      Left            =   15975
      TabIndex        =   69
      Top             =   4530
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame frm_Week 
      BackColor       =   &H00FFFFFF&
      Caption         =   " ���� ���� ���� "
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
      Height          =   855
      Left            =   8190
      TabIndex        =   37
      ToolTipText     =   "�ش� ���Ͽ��� ���డ���մϴ�"
      Top             =   7920
      Width           =   7155
      Begin VB.CheckBox chk_Week 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   6
         Left            =   5790
         TabIndex        =   26
         Top             =   390
         Value           =   1  'Ȯ��
         Width           =   615
      End
      Begin VB.CheckBox chk_Week 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   5
         Left            =   4890
         TabIndex        =   25
         Top             =   390
         Value           =   1  'Ȯ��
         Width           =   615
      End
      Begin VB.CheckBox chk_Week 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   4
         Left            =   4005
         TabIndex        =   24
         Top             =   390
         Value           =   1  'Ȯ��
         Width           =   615
      End
      Begin VB.CheckBox chk_Week 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   3
         Left            =   3105
         TabIndex        =   23
         Top             =   390
         Value           =   1  'Ȯ��
         Width           =   615
      End
      Begin VB.CheckBox chk_Week 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   2
         Left            =   2205
         TabIndex        =   22
         Top             =   390
         Value           =   1  'Ȯ��
         Width           =   615
      End
      Begin VB.CheckBox chk_Week 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   1
         Left            =   1320
         TabIndex        =   21
         Top             =   390
         Value           =   1  'Ȯ��
         Width           =   615
      End
      Begin VB.CheckBox chk_Week 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   0
         Left            =   420
         TabIndex        =   20
         Top             =   390
         Value           =   1  'Ȯ��
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   " [10,5,2]���� ���� "
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
      Height          =   915
      Left            =   8190
      TabIndex        =   67
      Top             =   6990
      Width           =   7155
      Begin VB.ComboBox cmb_Rotation_YN 
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
         ItemData        =   "FrmReg.frx":4415
         Left            =   330
         List            =   "FrmReg.frx":4417
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   68
         Top             =   390
         Width           =   2325
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   14490
      Top             =   105
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   " �����˻�"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   11.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Left            =   15
      TabIndex        =   56
      Top             =   6990
      Width           =   8130
      Begin VB.ComboBox cmb_GubunSrch 
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
         ItemData        =   "FrmReg.frx":4419
         Left            =   2130
         List            =   "FrmReg.frx":441B
         TabIndex        =   71
         Text            =   "cmb_Gubun"
         Top             =   600
         Width           =   3060
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
         Left            =   2100
         TabIndex        =   14
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
         Left            =   3945
         TabIndex        =   15
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
         IMEMode         =   10  '�ѱ� 
         Left            =   2070
         TabIndex        =   13
         Top             =   540
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
         ItemData        =   "FrmReg.frx":441D
         Left            =   480
         List            =   "FrmReg.frx":441F
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   12
         Top             =   540
         Width           =   1500
      End
      Begin Threed.SSCommand cmd_Search 
         Height          =   705
         Left            =   6105
         TabIndex        =   16
         Top             =   570
         Width           =   1185
         _Version        =   65536
         _ExtentX        =   2090
         _ExtentY        =   1244
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
         Picture         =   "FrmReg.frx":4421
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
         Left            =   3435
         TabIndex        =   58
         Top             =   1125
         Visible         =   0   'False
         Width           =   465
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
         Left            =   5280
         TabIndex        =   57
         Top             =   1125
         Visible         =   0   'False
         Width           =   465
      End
   End
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
      Height          =   960
      Left            =   15
      TabIndex        =   38
      Top             =   11940
      Width           =   15345
   End
   Begin VB.Frame frm_Rotation 
      Appearance      =   0  '���
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   15975
      TabIndex        =   32
      Top             =   6690
      Visible         =   0   'False
      Width           =   7185
      Begin VB.OptionButton Opt_Rotation 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00000000&
         Height          =   345
         Index           =   3
         Left            =   5550
         TabIndex        =   36
         Top             =   360
         Width           =   1305
      End
      Begin VB.OptionButton Opt_Rotation 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00000000&
         Height          =   345
         Index           =   2
         Left            =   3900
         TabIndex        =   35
         Top             =   360
         Width           =   1305
      End
      Begin VB.OptionButton Opt_Rotation 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00000000&
         Height          =   345
         Index           =   1
         Left            =   2250
         TabIndex        =   34
         Top             =   360
         Width           =   1305
      End
      Begin VB.OptionButton Opt_Rotation 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00000000&
         Height          =   345
         Index           =   0
         Left            =   600
         TabIndex        =   33
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
      ItemData        =   "FrmReg.frx":4772
      Left            =   15975
      List            =   "FrmReg.frx":4774
      TabIndex        =   31
      Text            =   "�˻�����"
      Top             =   6180
      Visible         =   0   'False
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
      Left            =   15975
      TabIndex        =   40
      Top             =   3195
      Visible         =   0   'False
      Width           =   2325
   End
   Begin ComctlLib.ListView ListView_REG 
      Height          =   5475
      Left            =   15
      TabIndex        =   19
      Top             =   1500
      Width           =   15330
      _ExtentX        =   27040
      _ExtentY        =   9657
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
      Left            =   14145
      TabIndex        =   18
      Top             =   765
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
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
      Picture         =   "FrmReg.frx":4776
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   585
      Index           =   5
      Left            =   13005
      TabIndex        =   17
      Top             =   765
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
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
      Picture         =   "FrmReg.frx":4AC7
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   540
      Index           =   6
      Left            =   18810
      TabIndex        =   27
      Top             =   6045
      Visible         =   0   'False
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
      Left            =   15975
      TabIndex        =   39
      Top             =   5490
      Visible         =   0   'False
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
      Height          =   3150
      Left            =   15
      TabIndex        =   41
      Top             =   8790
      Width           =   15330
      Begin VB.CheckBox chk_Lane 
         BackColor       =   &H00FFFFFF&
         Caption         =   "����6"
         Height          =   375
         Index           =   5
         Left            =   12960
         TabIndex        =   66
         Top             =   2040
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.CheckBox chk_Lane 
         BackColor       =   &H00FFFFFF&
         Caption         =   "����5"
         Height          =   375
         Index           =   4
         Left            =   11130
         TabIndex        =   65
         Top             =   2040
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.CheckBox chk_Lane 
         BackColor       =   &H00FFFFFF&
         Caption         =   "����4"
         Height          =   375
         Index           =   3
         Left            =   9300
         TabIndex        =   64
         Top             =   2040
         Width           =   1755
      End
      Begin VB.CheckBox chk_Lane 
         BackColor       =   &H00FFFFFF&
         Caption         =   "����3"
         Height          =   375
         Index           =   2
         Left            =   12960
         TabIndex        =   63
         Top             =   1770
         Width           =   1755
      End
      Begin VB.CheckBox chk_Lane 
         BackColor       =   &H00FFFFFF&
         Caption         =   "����2"
         Height          =   375
         Index           =   1
         Left            =   11130
         TabIndex        =   62
         Top             =   1770
         Width           =   1755
      End
      Begin VB.CheckBox chk_Lane 
         BackColor       =   &H00FFFFFF&
         Caption         =   "����1"
         Height          =   375
         Index           =   0
         Left            =   9300
         TabIndex        =   61
         Top             =   1770
         Width           =   1755
      End
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
         ItemData        =   "FrmReg.frx":4E18
         Left            =   9300
         List            =   "FrmReg.frx":4E22
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   11
         Top             =   1320
         Width           =   2070
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
         Left            =   7905
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   43
         Top             =   2220
         Visible         =   0   'False
         Width           =   1185
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
         ItemData        =   "FrmReg.frx":4E34
         Left            =   9300
         List            =   "FrmReg.frx":4E36
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   9
         Top             =   435
         Width           =   2070
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
         IMEMode         =   10  '�ѱ� 
         Left            =   1275
         MaxLength       =   12
         TabIndex        =   0
         Top             =   885
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
         IMEMode         =   10  '�ѱ� 
         Left            =   9300
         MaxLength       =   64
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   825
         Width           =   5220
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
         IMEMode         =   10  '�ѱ� 
         Left            =   5160
         MaxLength       =   15
         TabIndex        =   6
         Top             =   1320
         Width           =   2430
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
         Left            =   1275
         MaxLength       =   15
         TabIndex        =   2
         Top             =   1755
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
         IMEMode         =   10  '�ѱ� 
         Left            =   1275
         MaxLength       =   15
         TabIndex        =   1
         Top             =   1320
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
         IMEMode         =   10  '�ѱ� 
         Left            =   1275
         MaxLength       =   15
         TabIndex        =   3
         Top             =   2205
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
         Left            =   1275
         TabIndex        =   42
         Top             =   270
         Width           =   2475
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
         ItemData        =   "FrmReg.frx":4E38
         Left            =   5160
         List            =   "FrmReg.frx":4E3A
         TabIndex        =   5
         Text            =   "cmb_Dong"
         Top             =   900
         Width           =   2430
      End
      Begin MSMask.MaskEdBox MaskEdBox_Start 
         Height          =   375
         Left            =   5160
         TabIndex        =   7
         Top             =   1770
         Width           =   1545
         _ExtentX        =   2725
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
      Begin MSMask.MaskEdBox MaskEdBox_End 
         Height          =   375
         Left            =   5160
         TabIndex        =   8
         Top             =   2220
         Width           =   1545
         _ExtentX        =   2725
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
         Left            =   5160
         TabIndex        =   4
         Top             =   435
         Width           =   2430
         _ExtentX        =   4286
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   7
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
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   375
         Left            =   6720
         TabIndex        =   73
         Top             =   1770
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
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
         Format          =   140378115
         UpDown          =   -1  'True
         CurrentDate     =   36927
      End
      Begin MSComCtl2.DTPicker DTPicker4 
         Height          =   375
         Left            =   6720
         TabIndex        =   74
         Top             =   2220
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
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
         Format          =   140378115
         UpDown          =   -1  'True
         CurrentDate     =   36927
      End
      Begin Threed.SSCommand cmd_Button 
         Height          =   540
         Index           =   2
         Left            =   14115
         TabIndex        =   85
         Top             =   2520
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
         Enabled         =   0   'False
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "FrmReg.frx":4E3C
      End
      Begin Threed.SSCommand cmd_Button 
         Height          =   540
         Index           =   4
         Left            =   13005
         TabIndex        =   86
         Top             =   2520
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
         Enabled         =   0   'False
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "FrmReg.frx":518D
      End
      Begin Threed.SSCommand cmd_Button 
         Height          =   540
         Index           =   1
         Left            =   11895
         TabIndex        =   87
         Top             =   2520
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
         Enabled         =   0   'False
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "FrmReg.frx":54DE
      End
      Begin Threed.SSCommand cmd_Button 
         Height          =   540
         Index           =   3
         Left            =   10785
         TabIndex        =   88
         Top             =   2520
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
         Enabled         =   0   'False
         BevelWidth      =   1
         RoundedCorners  =   0   'False
         Picture         =   "FrmReg.frx":582F
      End
      Begin VB.Label Label8 
         BackStyle       =   0  '����
         Caption         =   "�������"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   12
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   8205
         TabIndex        =   60
         Top             =   1785
         Width           =   1065
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
         Height          =   315
         Left            =   8205
         TabIndex        =   59
         Top             =   1350
         Width           =   1065
      End
      Begin VB.Label Label4 
         BackStyle       =   0  '����
         Caption         =   "��       ��"
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
         Height          =   315
         Left            =   8205
         TabIndex        =   55
         Top             =   480
         Width           =   1065
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
         Left            =   4200
         TabIndex        =   54
         Top             =   480
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
         Left            =   3960
         TabIndex        =   53
         Top             =   915
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
         Height          =   315
         Index           =   0
         Left            =   165
         TabIndex        =   52
         Top             =   2205
         Width           =   1065
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
         Height          =   315
         Left            =   165
         TabIndex        =   51
         Top             =   1755
         Width           =   1065
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
         Left            =   4200
         TabIndex        =   50
         Top             =   1785
         Width           =   960
      End
      Begin VB.Label lbl_Object 
         BackStyle       =   0  '����
         Caption         =   "��       ��"
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
         Height          =   315
         Left            =   8205
         TabIndex        =   49
         Top             =   915
         Width           =   1065
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
         Left            =   4200
         TabIndex        =   48
         Top             =   2220
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
         Left            =   3960
         TabIndex        =   47
         Top             =   1350
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
         Height          =   315
         Left            =   165
         TabIndex        =   46
         Top             =   450
         Width           =   1065
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
         Height          =   315
         Left            =   165
         TabIndex        =   45
         Top             =   1305
         Width           =   1065
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
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   165
         TabIndex        =   44
         Top             =   885
         Width           =   1065
      End
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   585
      Index           =   8
      Left            =   11190
      TabIndex        =   72
      Top             =   765
      Width           =   1755
      _Version        =   65536
      _ExtentX        =   3096
      _ExtentY        =   1032
      _StockProps     =   78
      Caption         =   "������̷���ȸ"
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
      Picture         =   "FrmReg.frx":5B80
   End
   Begin VB.Label lbl_App 
      BackStyle       =   0  '����
      Caption         =   "�����������"
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
      Height          =   315
      Left            =   16065
      TabIndex        =   84
      ToolTipText     =   "����Ͼ� üũ������ ���, ��й�ȣ �ʱ�ȭ�ǹǷ� �ݵ�� ���ѳ��� �����ؾ��մϴ�."
      Top             =   8985
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label lbl_gubun_tmp 
      BackColor       =   &H0080C0FF&
      Caption         =   "lbl_gubun_tmp"
      Height          =   375
      Left            =   15975
      TabIndex        =   81
      Top             =   2790
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lbl_object_tmp 
      BackColor       =   &H0080C0FF&
      Caption         =   "lbl_object_tmp"
      Height          =   375
      Left            =   15975
      TabIndex        =   80
      Top             =   2370
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lbl_ho_tmp 
      BackColor       =   &H0080C0FF&
      Caption         =   "lbl_ho_tmp"
      Height          =   375
      Left            =   15975
      TabIndex        =   79
      Top             =   1950
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lbl_dong_tmp 
      BackColor       =   &H0080C0FF&
      Caption         =   "lbl_dong_tmp"
      Height          =   375
      Left            =   15975
      TabIndex        =   78
      Top             =   1530
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lbl_Carno_tmp 
      BackColor       =   &H0080C0FF&
      Caption         =   "lbl_Carno_tmp"
      Height          =   375
      Left            =   15975
      TabIndex        =   77
      Top             =   270
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lbl_Name_tmp 
      BackColor       =   &H0080C0FF&
      Caption         =   "lbl_Name_tmp"
      Height          =   375
      Left            =   15975
      TabIndex        =   76
      Top             =   690
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lbl_CarModel_tmp 
      BackColor       =   &H0080C0FF&
      Caption         =   "lbl_CarModel_tmp"
      Height          =   375
      Left            =   15975
      TabIndex        =   75
      Top             =   1110
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lbl_title 
      BackColor       =   &H00404040&
      Caption         =   "���� ��� ����"
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
      TabIndex        =   28
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
      TabIndex        =   30
      Top             =   1005
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
      TabIndex        =   29
      Top             =   1005
      Width           =   900
   End
End
Attribute VB_Name = "FrmReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CAR_NO_TMP As String
Dim PART_NAME_TMP As String
Dim RegQry As String
Dim APP_INIT_PASSWORD As String

Private Sub cmb_GB_Click()
    If (cmb_GB.text = "�Ҽ�/����" Or cmb_GB.text = "��/ȣ��") Then
        txt_tmpCarNo.Enabled = False
        txt_tmpCarNo.Visible = False
        
        cmb_GubunSrch.Visible = False
        cmb_GubunSrch.Enabled = False
        
        Label3.Visible = True
        Label6.Visible = True
        cmbDong.Enabled = True
        cmbDong.Visible = True
        cmbHo.Enabled = True
        cmbHo.Visible = True
        
    ElseIf (cmb_GB.text = "�� ��") Then
        
        txt_tmpCarNo.Enabled = False
        txt_tmpCarNo.Visible = False

        Label3.Visible = False
        Label6.Visible = False
        cmbDong.Enabled = False
        cmbDong.Visible = False
        cmbHo.Enabled = False
        cmbHo.Visible = False
        
        cmb_GubunSrch.Visible = True
        cmb_GubunSrch.Enabled = True
    
    Else
        txt_tmpCarNo.Enabled = True
        txt_tmpCarNo.Visible = True
        
        cmb_GubunSrch.Visible = False
        cmb_GubunSrch.Enabled = False
        
        Label3.Visible = False
        Label6.Visible = False
        cmbDong.Enabled = False
        cmbDong.Visible = False
        cmbHo.Enabled = False
        cmbHo.Visible = False
    End If
End Sub

Private Sub cmd_PWInit_Click()
    On Error GoTo Err_p
    
    adoConn.Execute "UPDATE tb_reg     SET APP_PW='" & APP_INIT_PASSWORD & "', APP_CERTIFY_DATE =Null WHERE CAR_NO = '" & txt_CarNo & "'"
    adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('" & Trim(txt_CarNo) & "', 'HOST','�ۺ�й�ȣ �ʱ�ȭ',''," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
    List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_CarNo & "   �� ��й�ȣ �ʱ�ȭ ����", 0
    Call DataLogger("[REG App PW Init]    " & txt_CarNo & "   �� ��й�ȣ �ʱ�ȭ ����")
    Exit Sub
    
Err_p:
    Call DebugLogger("[REG App PW Init]    " & Err.Description)
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim rs As Recordset
Dim qry As String
Dim bQryResult As Boolean


    Left = (Screen.width - width) / 2   ' ���� ���η� �߾ӿ� �����ϴ�.
    Top = (Screen.height - height) / 2   ' ���� ���η� �߾ӿ� �����ϴ�.

    cmbDong.Enabled = True
    cmbDong.Visible = True
    Label3.Enabled = True
    Label3.Visible = True
    cmbHo.Enabled = True
    cmbHo.Visible = True
    Label6.Enabled = True
    Label6.Visible = True
    cmd_PWInit.Enabled = False
    
    'Me.cmb_Gubun = Me.cmb_Gubun.List(0)
    RegQry = "SELECT * From tb_reg ORDER BY CAR_GUBUN ASC, DRIVER_DEPT ASC, DRIVER_CLASS ASC"

    If (Glo_User_Type = "����1/����2") Then
        With cmb_Gubun
            .AddItem "�����"
            .AddItem "������"
            .AddItem "���¾�ü"
            .AddItem "����ó��"
            .AddItem "��������"
            .text = cmb_Gubun.List(0)
        End With
        'lbl_dept(0).Caption = "��  �� :"
        lbl_dept(2).Caption = "    ��     ��"
        lbl_dept(3).Caption = "    ��     ��"
        With cmb_Search
            .AddItem "��ü"
            .AddItem "�����"
            .AddItem "������"
            .AddItem "���¾�ü"
            .AddItem "����ó��"
            .AddItem "��������"
            .AddItem "�Ⱓ�ʰ�"
            .text = cmb_Search.List(0)
        End With
    Else
        With cmb_Gubun
            .AddItem "���ֹ�"
            .AddItem "��"
            .AddItem "������"
            .AddItem "���¾�ü"
            .AddItem "����ó��"
            .AddItem "��������"
            .text = cmb_Gubun.List(0)
        End With
        'lbl_dept(0).Caption = "���� �� :"
        'lbl_dept(1).Caption = "���� ȣ :"
        lbl_dept(2).Caption = "����  ��"
        lbl_dept(3).Caption = "����  ȣ"
        With cmb_Search
            .AddItem "��ü"
            .AddItem "���ֹ�"
            .AddItem "��"
            .AddItem "������"
            .AddItem "���¾�ü"
            '.AddItem "�ӽõ��"
            .AddItem "����ó��"
            .AddItem "��������"
            .AddItem "�Ⱓ�ʰ�"
            .text = cmb_Search.List(0)
        End With
    End If
    
    For i = 1 To MAX_REG_GUBUN
        If (Glo_RegGubun(i) <> "") Then
            cmb_Gubun.AddItem Glo_RegGubun(i)
        End If
    Next
    
    If (cmb_Gubun.ListCount > 0) Then
        cmb_Gubun.text = cmb_Gubun.List(0)
    Else
        cmb_Gubun.text = ""
    End If
        
        
        
        
    
    
        If (Glo_User_Type = "����1/����2") Then
            cmb_GB.AddItem "������ȣ"
            cmb_GB.AddItem "�� ��"
            cmb_GB.AddItem "�Ҽ�/����"
            cmb_GB.AddItem "�� ��"
    
            txt_tmpCarNo.Enabled = True
            txt_tmpCarNo.Visible = True
    
            Label3.Caption = "�Ҽ�"
            Label6.Caption = "����"
            Label3.Visible = False
            Label6.Visible = False
            cmbDong.Enabled = False
            cmbDong.Visible = False
            cmbHo.Enabled = False
            cmbHo.Visible = False
        Else
            cmb_GB.AddItem "������ȣ"
            cmb_GB.AddItem "�� ��"
            cmb_GB.AddItem "��/ȣ��"
            cmb_GB.AddItem "�� ��"
    
            txt_tmpCarNo.Enabled = True
            txt_tmpCarNo.Visible = True
    
            Label3.Caption = "��"
            Label6.Caption = "ȣ��"
            Label3.Visible = False
            Label6.Visible = False
            cmbDong.Enabled = False
            cmbDong.Visible = False
            cmbHo.Enabled = False
            cmbHo.Visible = False
        End If
        

        
        
        
    
    
        '����ǰ��� ��ư
        cmd_Button(7).Enabled = False
        cmd_Button(7).Visible = False
        '��������
'        Label5.Enabled = False
'        Label5.Visible = False
'        cmb_Rotation.Enabled = False
'        cmb_Rotation.Visible = False
        '���ϼ���
        
        
        '10,5,2 ���� ���� ����
        With cmb_Rotation_YN
            .AddItem "����"
            .AddItem "������"
            .text = cmb_Rotation_YN.List(1)
        End With
        If (Glo_ROTATION = "������") Then
            Frame3.Enabled = False
            Frame3.Visible = False
            'cmb_Rotation_YN.Enabled = False
        Else
            Frame3.Enabled = True
            Frame3.Visible = True
            'cmb_Rotation_YN.Enabled = True
        End If
        
        
        '���� ������� ����
        frm_Week.Visible = True
        For i = 0 To 6
            If (Glo_WEEK_YN = "Y") Then
                frm_Week.Enabled = True
                frm_Week.Visible = True
                chk_Week(i).Enabled = True
                chk_Week(i).Visible = True
            Else
                frm_Week.Enabled = False
                frm_Week.Visible = False
                chk_Week(i).Enabled = False
                chk_Week(i).Visible = True
            End If
        Next
        chk_Week(5).value = 0
        chk_Week(6).value = 0
        
        '����
    '    cmd_Button(5).Enabled = False
    '    cmd_Button(5).Visible = False
    
        
        If (LANE1_YN = "Y") Then
            chk_Lane(0).Caption = LANE1_Name
            chk_Lane(0).value = 1
        Else
            chk_Lane(0).Caption = "�̻��"
            chk_Lane(0).Enabled = False
            chk_Lane(0).value = 0
        End If
        If (LANE2_YN = "Y") Then
            chk_Lane(1).Caption = LANE2_Name
            chk_Lane(1).value = 1
        Else
            chk_Lane(1).Caption = "�̻��"
            chk_Lane(1).Enabled = False
            chk_Lane(1).value = 0
        End If
        If (LANE3_YN = "Y") Then
            chk_Lane(2).Caption = LANE3_Name
            chk_Lane(2).value = 1
        Else
            chk_Lane(2).Caption = "�̻��"
            chk_Lane(2).Enabled = False
            chk_Lane(2).value = 0
        End If
        If (LANE4_YN = "Y") Then
            chk_Lane(3).Caption = LANE4_Name
            chk_Lane(3).value = 1
        Else
            chk_Lane(3).Caption = "�̻��"
            chk_Lane(3).Enabled = False
            chk_Lane(3).value = 0
        End If
        If (LANE5_YN = "Y") Then
            chk_Lane(4).Caption = LANE5_Name
            chk_Lane(4).value = 1
        Else
            chk_Lane(4).Caption = "�̻��"
            chk_Lane(4).Enabled = False
            chk_Lane(4).value = 0
        End If
        If (LANE6_YN = "Y") Then
            chk_Lane(5).Caption = LANE6_Name
            chk_Lane(5).value = 1
        Else
            chk_Lane(5).Caption = "�̻��"
            chk_Lane(5).Enabled = False
            chk_Lane(5).value = 0
        End If
        
        chk_Lane(0).Visible = False
        chk_Lane(1).Visible = False
        chk_Lane(2).Visible = False
        chk_Lane(3).Visible = False
        chk_Lane(4).Visible = False
        chk_Lane(5).Visible = False
        If (Glo_Screen_No >= 1) Then
            chk_Lane(0).Visible = True
        End If
        If (Glo_Screen_No >= 2) Then
            chk_Lane(1).Visible = True
        End If
        
        If (Glo_Screen_No >= 4) Then
            chk_Lane(2).Visible = True
            chk_Lane(3).Visible = True
        End If
        If (Glo_Screen_No >= 6) Then
            chk_Lane(4).Visible = True
            chk_Lane(5).Visible = True
        End If
        
        
        
        If (Glo_RegMonFee_YN = "Y") Then
            Label1.Caption = "��     ��"
            MaskEdBox_Fee.Visible = True
            Label1.Visible = True
        Else
            Label1.Caption = "..."
            MaskEdBox_Fee.Visible = False
            Label1.Visible = False
        End If


   ' End If
   
   
    If (Able_WebDC = False) Then
        lbl_App.Visible = False
        chk_App.Visible = False
        cmd_PWInit.Visible = False
    Else
        lbl_App.Visible = True
        chk_App.Visible = True
        cmd_PWInit.Visible = True
    End If
    
    Call Clear_Field
    Call ListView_REG_Draw
    Call ListView_REG_SQL
    
    cmb_GB.ListIndex = 0
    chk_App.value = 1
  
    
    List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    �������/���� ����...!!", 0
    Call DataLogger("[REG Formload]    " & "�������/���� ����...!!")
    'Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    �������/���� ����...!!")
    
    Call SaveReg2
    List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    ������� ����", 0
    Call DataLogger("[FrmLoad]    " & "������� ����")
    
    
    APP_INIT_PASSWORD = EncodeNDE01("12345678", "www.jawootek.com") '��ȣȭ

End Sub


'1���� ����
Private Sub cmd_Month_Click()
    
    If (MaskEdBox_End.text <> "9999-12-31") Then
        MaskEdBox_End.text = DateAdd("m", 1, MaskEdBox_End.text)
    End If

End Sub

Public Sub ListView_REG_SQL()
    Dim bQryResult As Boolean
    Dim rs As Recordset
    Dim qry As String
    Dim itmX As ListItem
    Dim INDEX_NO As Long
    Dim i As Integer
    Dim AppYN As Boolean

    AppYN = Able_WebDC

    'cmbDong
    Call Set_cmbDong
    'cmbHo
    Call Set_cmbHo
    
    Call Set_cmbGubunSrch
    
    INDEX_NO = 1
    Set rs = New ADODB.Recordset
    'rs.Open RegQry, adoConn
    bQryResult = DataBaseQuery(rs, adoConn, RegQry, False)
    If (bQryResult = False) Then
        List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    ��Ʈ��ũ �� DB ���˹ٶ��ϴ�", 0
        Call DataLogger("[FrmReg]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
        Exit Sub
    End If
    
    
    lbl_COUNT = rs.RecordCount
    Do While Not (rs.EOF)
        Set itmX = ListView_REG.ListItems.Add(, , "" & INDEX_NO)
        
        i = 1
        itmX.SubItems(i) = "" & rs!CAR_NO: i = i + 1
        itmX.SubItems(i) = "" & rs!CAR_MODEL: i = i + 1
        itmX.SubItems(i) = "" & rs!CAR_GUBUN: i = i + 1
        
        If (Glo_Screen_No >= 1) Then
            If (rs!LANE1 = "Y") Then
                itmX.SubItems(i) = "" & "Y": i = i + 1
            Else
                itmX.SubItems(i) = "" & "N": i = i + 1
            End If
        End If
        
        If (Glo_Screen_No >= 2) Then
            If (rs!LANE2 = "Y") Then
                itmX.SubItems(i) = "" & "Y": i = i + 1
            Else
                itmX.SubItems(i) = "" & "N": i = i + 1
            End If
        End If
        
        If (Glo_Screen_No >= 4) Then
            If (rs!LANE3 = "Y") Then
                itmX.SubItems(i) = "" & "Y": i = i + 1
            Else
                itmX.SubItems(i) = "" & "N": i = i + 1
            End If
            If (rs!LANE4 = "Y") Then
                itmX.SubItems(i) = "" & "Y": i = i + 1
            Else
                itmX.SubItems(i) = "" & "N": i = i + 1
            End If
        End If
    
        If (Glo_Screen_No >= 6) Then
            If (rs!LANE5 = "Y") Then
                itmX.SubItems(i) = "" & "Y": i = i + 1
            Else
                itmX.SubItems(i) = "" & "N": i = i + 1
            End If
            If (rs!LANE6 = "Y") Then
                itmX.SubItems(i) = "" & "Y": i = i + 1
            Else
                itmX.SubItems(i) = "" & "N": i = i + 1
            End If
        End If
        
        
        If (Glo_ROTATION <> "������") Then
            If (rs!Rotation = "Y") Then
                    itmX.SubItems(i) = "" & "����": i = i + 1
            Else
                    itmX.SubItems(i) = "" & "������": i = i + 1
            End If
        End If
    
        If (Glo_RegMonFee_YN = "Y") Then
            itmX.SubItems(i) = "" & rs!CAR_FEE: i = i + 1
        End If
        itmX.SubItems(i) = "" & rs!DRIVER_NAME: i = i + 1
        itmX.SubItems(i) = "" & rs!DRIVER_PHONE: i = i + 1
        itmX.SubItems(i) = "" & rs!DRIVER_DEPT: i = i + 1
        itmX.SubItems(i) = "" & rs!DRIVER_CLASS: i = i + 1
        'itmX.SubItems(i) = "" & rs!Start_Date: i = i + 1
        'itmX.SubItems(i) = "" & rs!End_Date: i = i + 1
        'itmX.SubItems(i) = "" & Left(rs!Start_Date, 8): i = i + 1
        'itmX.SubItems(i) = "" & Left(rs!End_Date, 8): i = i + 1
        'itmX.SubItems(i) = "" & rs!REG_DATE: i = i + 1
        itmX.SubItems(i) = "" & Left(rs!START_DATE, 10): i = i + 1
        itmX.SubItems(i) = "" & Left(rs!END_DATE, 10): i = i + 1
        itmX.SubItems(i) = "" & Format(rs!REG_DATE, "yyyy-mm-dd hh:nn:ss"): i = i + 1
        itmX.SubItems(i) = "" & rs!Update_date: i = i + 1
        If (Glo_RegMonFee_YN = "Y") Then
            itmX.SubItems(i) = "" & rs!FEE_DATE: i = i + 1
        End If
        itmX.SubItems(i) = "" & rs!DAY_ROTATION_YN: i = i + 1
        itmX.SubItems(i) = "" & rs!REG_PART: i = i + 1
        itmX.SubItems(i) = "" & rs!ETC: i = i + 1
        
        If (Glo_WEEK_YN = "Y") Then
            itmX.SubItems(i) = "" & rs!WEEK1: i = i + 1
            itmX.SubItems(i) = "" & rs!WEEK2: i = i + 1
            itmX.SubItems(i) = "" & rs!WEEK3: i = i + 1
            itmX.SubItems(i) = "" & rs!WEEK4: i = i + 1
            itmX.SubItems(i) = "" & rs!WEEK5: i = i + 1
            itmX.SubItems(i) = "" & rs!WEEK6: i = i + 1
            itmX.SubItems(i) = "" & rs!WEEK7: i = i + 1
        End If
        
'        If (AppYN = True) Then
'            itmX.SubItems(i) = "" & rs!APP_YN: i = i + 1
'        End If
        
        rs.MoveNext
        INDEX_NO = INDEX_NO + 1
    Loop
    Set rs = Nothing
End Sub

Public Sub ListView_REG_Draw()
Dim Column_to_size As Integer
Dim AppYN As Boolean

AppYN = Able_WebDC
    

With Me
    Call ListViewExtended(.ListView_REG)
    .ListView_REG.View = lvwReport
    .ListView_REG.ListItems.Clear
    .ListView_REG.ColumnHeaders.Clear
    .ListView_REG.ColumnHeaders.Add , , " No  "
    .ListView_REG.ColumnHeaders.Add , , " ������ȣ        "
    .ListView_REG.ColumnHeaders.Add , , " ������     "
    .ListView_REG.ColumnHeaders.Add , , " ��������   "
    
    
    If (Glo_Screen_No >= 1) Then
        .ListView_REG.ColumnHeaders.Add , , LANE1_Name
    End If
    If (Glo_Screen_No >= 2) Then
        .ListView_REG.ColumnHeaders.Add , , LANE2_Name
    End If
    If (Glo_Screen_No >= 4) Then
        .ListView_REG.ColumnHeaders.Add , , LANE3_Name
        .ListView_REG.ColumnHeaders.Add , , LANE4_Name
    End If
    If (Glo_Screen_No >= 6) Then
        .ListView_REG.ColumnHeaders.Add , , LANE5_Name
        .ListView_REG.ColumnHeaders.Add , , LANE6_Name
    End If
    
    If (Glo_ROTATION <> "������") Then
        .ListView_REG.ColumnHeaders.Add , , "��������"
    End If
    
    If (Glo_RegMonFee_YN = "Y") Then
        .ListView_REG.ColumnHeaders.Add , , " �������   "
    End If
    .ListView_REG.ColumnHeaders.Add , , " ��    ��     "
    .ListView_REG.ColumnHeaders.Add , , " �� �� ó              "
    If (Glo_User_Type = "����1/����2") Then
        ListView_REG.ColumnHeaders.Add , , " ��    ��    "
        ListView_REG.ColumnHeaders.Add , , " ��    ��    "
    Else
        ListView_REG.ColumnHeaders.Add , , " ����  ��    "
        ListView_REG.ColumnHeaders.Add , , " ����  ȣ    "
    End If
    .ListView_REG.ColumnHeaders.Add , , " �� �� ��        "
    .ListView_REG.ColumnHeaders.Add , , " �� �� ��        "
    .ListView_REG.ColumnHeaders.Add , , " �� �� ��                       "
    .ListView_REG.ColumnHeaders.Add , , " �� �� ��                       "
    If (Glo_RegMonFee_YN = "Y") Then
        .ListView_REG.ColumnHeaders.Add , , " �� �� ��   "
    End If
    .ListView_REG.ColumnHeaders.Add , , " �����뺸 "
    .ListView_REG.ColumnHeaders.Add , , " ��� "
    .ListView_REG.ColumnHeaders.Add , , " ��Ÿ "
    
    If (Glo_WEEK_YN = "Y") Then
        .ListView_REG.ColumnHeaders.Add , , " �� "
        .ListView_REG.ColumnHeaders.Add , , " ȭ "
        .ListView_REG.ColumnHeaders.Add , , " �� "
        .ListView_REG.ColumnHeaders.Add , , " �� "
        .ListView_REG.ColumnHeaders.Add , , " �� "
        .ListView_REG.ColumnHeaders.Add , , " �� "
        .ListView_REG.ColumnHeaders.Add , , " �� "
    End If

'    If (AppYN = True) Then
'        .ListView_REG.ColumnHeaders.Add , , " ������ �����   "
'    End If
    
    For Column_to_size = 0 To .ListView_REG.ColumnHeaders.Count - 2
         SendMessage .ListView_REG.hwnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
    Next
End With

End Sub



Private Sub Label1_Click()
    If (Label1.Caption = "...") Then
        Label1.Caption = "��     ��"
        MaskEdBox_Fee.Visible = True
        Glo_RegMonFee_YN = "Y"
        Call Put_Ini("System Config", "RegMonFee_YN", Glo_RegMonFee_YN)
    Else
        Label1.Caption = "..."
        MaskEdBox_Fee.Visible = False
        Glo_RegMonFee_YN = "N"
        Call Put_Ini("System Config", "RegMonFee_YN", Glo_RegMonFee_YN)
    End If

    Msg_Box.Label1 = "������ϰ��� ���� ����ƽ��ϴ�" & vbCrLf & vbCrLf & "�ϰ�����Ϸ��� ����������" & vbCrLf & "�ٽ� �Ͻʽÿ�"
    Msg_Box.Show 1
    Call Clear_Field
    Call ListView_REG_Draw
    Call ListView_REG_SQL
    
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
    On Error Resume Next
    
    cmd_PWInit.Enabled = True
    ListView_REG.SetFocus
    txt_CarNo = ListView_REG.SelectedItem.SubItems(1)       '������ȣ
    lbl_gubun_tmp = ListView_REG.SelectedItem.SubItems(3)   '����
    
    lbl_Carno_tmp = txt_CarNo
    lbl_Name_tmp = txt_Name
    lbl_CarModel_tmp = txt_CarModel
    lbl_dong_tmp = cmb_Dong.text
    lbl_ho_tmp = txt_Ho
    lbl_object_tmp = txt_Object
    
    cmd_Button(3).Enabled = True
    cmd_Button(1).Enabled = False
    cmd_Button(4).Enabled = True
    cmd_Button(2).Enabled = True
    
    Call Search_Record
End Sub

Public Sub Clear_Field()
    cmd_Button(3).Enabled = True    '�ʱ�ȭ
    cmd_Button(1).Enabled = True    '���
    cmd_Button(4).Enabled = False   '����
    cmd_Button(2).Enabled = False   '����

   
    CAR_NO_TMP = ""
    txt_Num.text = ""
    txt_CarNo.text = ""
    txt_Name.text = ""
    txt_Phone.text = ""
    txt_CarModel.text = ""
    cmb_Gubun.ListIndex = 0
    
    If (Glo_User_Type = "����1/����2") Then
        cmb_Rotation.ListIndex = 1
    Else
        cmb_Rotation.ListIndex = 0
    End If
    
    cmb_Rotation_YN.ListIndex = 1
    'txt_Dong.Text = ""
    cmb_Dong.text = ""
    txt_Ho.text = ""
    
    lbl_Carno_tmp = ""
    lbl_Name_tmp = ""
    lbl_CarModel_tmp = ""
    lbl_dong_tmp = ""
    lbl_ho_tmp = ""
    lbl_object_tmp = ""
    lbl_gubun_tmp = ""
    
    MaskEdBox_Start.text = Format(Now, "yyyy-mm-dd")
    '������ ����
    Select Case Glo_EndDate
        Case 99
            MaskEdBox_End.text = "9999-12-31"
        Case Else
            MaskEdBox_End.text = Format(DateAdd("m", Glo_EndDate, Date), "yyyy-mm-dd")
    End Select

    DTPicker3.Format = dtpCustom
    DTPicker3.CustomFormat = "HH:mm"
    DTPicker3.Refresh
    DTPicker3.value = Format("00:00")
    
    DTPicker4.Format = dtpCustom
    DTPicker4.CustomFormat = "HH:mm"
    DTPicker4.Refresh
    DTPicker4.value = Format("23:59")
    
    
    MaskEdBox_Fee.text = "0"
    txt_Object.text = ""
    chk_Week(0).value = 1
    chk_Week(1).value = 1
    chk_Week(2).value = 1
    chk_Week(3).value = 1
    chk_Week(4).value = 1
    chk_Week(5).value = 1
    chk_Week(6).value = 1
    
    On Error Resume Next
    txt_CarNo.SetFocus
    cmd_PWInit.Enabled = False
    'chk_App.value = False
    
End Sub

'������ ����
Sub Delete_Record()
    Dim tmpLane1, tmpLane2, tmpLane3, tmpLane4, tmpLane5, tmpLane6 As String
    Dim tmpWeek1, tmpWeek2, tmpWeek3, tmpWeek4, tmpWeek5, tmpWeek6, tmpWeek7 As String
    Dim tmpCarNo, tmpName, tmpCarModel, tmpObject, tmpDong, tmpHo, stDate, edDate As String
    Dim tmpPhone, tmpGubun, tmpRegDate, tmpUpdate, tmpFeeDate, tmpRegPart, tmpAction, tmpAfterCarNo, tmpActionID As String
    Dim tmpDayRot, tmpRotation As String
    Dim tmpFee As Long
    Dim sApp, sAppPW, sApp_YesDate, sApp_Cert_Date, sLog_data As String
    
    Dim sQry As String
    Dim bQryResult As Boolean
    Dim rs As Recordset
    
    
On Error GoTo Err_p
    
    Dim sSaveTableName As String
    sSaveTableName = "tb_reg"
    
    sQry = "SELECT * from " & sSaveTableName & " WHERE CAR_NO = '" & txt_CarNo & "'"
    Set rs = New ADODB.Recordset
     bQryResult = DataBaseQuery(rs, adoConn, sQry, False)
     If (bQryResult = False) Then
        List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    ��Ʈ��ũ �� DB ���˹ٶ��ϴ�", 0
        Call DataLogger("[FrmReg Delete Record]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
        Exit Sub
    End If
    
    
    If (Not rs.EOF) Then
        tmpCarNo = rs!CAR_NO
        
        tmpCarModel = rs!CAR_MODEL
        tmpGubun = rs!CAR_GUBUN
        tmpFee = rs!CAR_FEE
        tmpName = rs!DRIVER_NAME
        tmpPhone = rs!DRIVER_PHONE
        tmpDong = rs!DRIVER_DEPT
        tmpHo = rs!DRIVER_CLASS
        stDate = Format(MaskEdBox_Start, "YYYYMMDD") & "000000"
        edDate = Format(MaskEdBox_End, "YYYYMMDD") & "235959"
        tmpObject = rs!ETC
        tmpRegDate = Format(Now, "yyyy-mm-dd hh:nn:ss")
        tmpUpdate = ""
        tmpFeeDate = ""
        tmpRegPart = Glo_PartName
        tmpAction = "����"
        tmpAfterCarNo = ""
        tmpActionID = Glo_Login_ID
        tmpDayRot = rs!DAY_ROTATION_YN
        tmpRotation = rs!Rotation
        tmpLane1 = rs!LANE1: tmpLane2 = rs!LANE2: tmpLane3 = rs!LANE3: tmpLane4 = rs!LANE4: tmpLane5 = rs!LANE5: tmpLane6 = rs!LANE6
        tmpWeek1 = rs!WEEK1: tmpWeek2 = rs!WEEK2: tmpWeek3 = rs!WEEK3: tmpWeek4 = rs!WEEK4: tmpWeek5 = rs!WEEK5: tmpWeek6 = rs!WEEK6: tmpWeek7 = rs!WEEK7:
        sApp = rs!APP_YN
        sAppPW = rs!APP_PW
        sApp_YesDate = rs!APP_YES_DATE
        sApp_Cert_Date = rs!APP_CERTIFY_DATE
        sLog_data = Format(Now, "yyyy-mm-dd hh:nn:ss")
    End If
    
    Set rs = Nothing
    
    
    
    
    

    sQry = "DELETE FROM " & sSaveTableName & " WHERE CAR_NO = '" & txt_CarNo & "'"
    bQryResult = DataBaseQueryExec(adoConn, sQry, NWERR_GATE_STAY)
    If (bQryResult = False) Then
        List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    ��Ʈ��ũ �� DB ���˹ٶ��ϴ�", 0
        Call DataLogger("[FrmReg]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
        Exit Sub
    End If

    'sQry = "INSERT INTO tb_reg_log VALUES ('" & txt_CarNo & "', '" & tmpCarModel & "', '" & cmb_Gubun.text & "', '" & MaskEdBox_Fee.text & "', '" & tmpName & "', '" & txt_Phone & "', '" & cmb_Dong & "', '" & txt_Ho & "', '" & stDate & "', '" & edDate & "', '" & tmpObject & "', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "', '', '', '" & cmb_Rotation.text & "', '" & Glo_PartName & "', '����', '', '" & Glo_Login_ID & "', '" & tmpLane1 & "', '" & tmpLane2 & "', '" & tmpLane3 & "', '" & tmpLane4 & "', '" & tmpLane5 & "', '" & tmpLane6 & "', '" & tmpWeek1 & "', '" & tmpWeek2 & "', '" & tmpWeek3 & "', '" & tmpWeek4 & "', '" & tmpWeek5 & "', '" & tmpWeek6 & "', '" & sChkWeek7 & "', '" & tmpDayRot & "', '" & sApp & "' )"
    sQry = "INSERT INTO tb_reg_log (CAR_NO, CAR_MODEL, CAR_GUBUN, CAR_FEE, DRIVER_NAME, DRIVER_PHONE, DRIVER_DEPT, DRIVER_CLASS, START_DATE, END_DATE, ETC, REG_DATE, UPDATE_DATE, FEE_DATE,DAY_ROTATION_YN,REG_PART,ACTION_LOG,AFTER_CAR_NO,ACTION_ID,  LANE1,LANE2,LANE3,LANE4,LANE5,LANE6,WEEK1,WEEK2,WEEK3,WEEK4,WEEK5,WEEK6,WEEK7,ROTATION,APP_YN,APP_PW,APP_YES_DATE,APP_CERTIFY_DATE,LOG_DATE) "
    sQry = sQry & " VALUES ('" & txt_CarNo & "', '" & tmpCarModel & "', '" & tmpGubun & "', '" & tmpFee & "', '" & tmpName & "', '" & tmpPhone & "', '" & tmpDong & "', '" & tmpHo & "', '" & stDate & "', '" & edDate & "', '" & tmpObject & "', '" & tmpRegDate & "', '', '', '" & tmpDayRot & "', '" & tmpRegPart & "', '" & tmpAction & "', '" & tmpAfterCarNo & "', '" & tmpActionID & "','" & tmpLane1 & "', '" & tmpLane2 & "', '" & tmpLane3 & "', '" & tmpLane4 & "', '" & tmpLane5 & "', '" & tmpLane6 & "', '" & tmpWeek1 & "', '" & tmpWeek2 & "', '" & tmpWeek3 & "', '" & tmpWeek4 & "', '" & tmpWeek5 & "', '" & tmpWeek6 & "', '" & tmpWeek7 & "','" & tmpRotation & "', '" & sApp & "', '" & sAppPW & "', '" & sApp_YesDate & "', '" & sApp_Cert_Date & "', '" & sLog_data & "')"
    bQryResult = DataBaseQueryExec(adoConn, sQry, NWERR_GATE_STAY)
    If (bQryResult = False) Then
        List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    ��Ʈ��ũ �� DB ���˹ٶ��ϴ�", 0
        Call DataLogger("[FrmReg]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
        Exit Sub
    End If


    ' �����湮���� ID ������� ���� : ���ֹ��� ���� �� ���� ����� ��� �� �Ϻ������� ������ ��쿡�� �����湮������ �� �־���ϹǷ�.
    

    List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_CarNo & "    �������� ���� �Ϸ�", 0
    Call DataLogger("[REG Button]    " & txt_CarNo & "    �������� ���� �Ϸ�")
    'Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_CarNo & "    �������� ���� �Ϸ�")
    Call ListView_REG_Draw
    Call ListView_REG_SQL
    
    Exit Sub
Err_p:
    List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & "[DELETE RECORD] " & Err.Description, 0
    Call DataLogger("[REG DELETE RECORD]    " & Err.Description)
End Sub

Sub Insert_Record()
    Dim rs_COUNT As Recordset
    Dim rs As Recordset
    Dim SQL_COUNT As String
    Dim SQL_QUARY As String
    Dim i As Integer
    Dim Cnt As Integer
    Dim tmp As String
    Dim tmpCarNo, tmpName, tmpPhone, tmpCarModel, tmpObject, tmpDong, tmpHo, stDate, edDate As String
    Dim P As String
    Dim sChkLane1 As String
    Dim sChkLane2 As String
    Dim sChkLane3 As String
    Dim sChkLane4 As String
    Dim sChkLane5 As String
    Dim sChkLane6 As String
    
    Dim sChkWeek1 As String
    Dim sChkWeek2 As String
    Dim sChkWeek3 As String
    Dim sChkWeek4 As String
    Dim sChkWeek5 As String
    Dim sChkWeek6 As String
    Dim sChkWeek7 As String
    
    Dim sRotation As String
    
    Dim sApp As String
    Dim sApp_YesDate As String
    Dim sApp_Certify_date As String
    Dim sLog_data As String
    
    Dim sTmp As String
    Dim sQry As String
    Dim bQryResult As Boolean
    
On Error GoTo Err_p

    
    If (Glo_Screen_No >= 1) Then
        If chk_Lane(0).value = 1 Then
            sChkLane1 = "Y"
        Else
            sChkLane1 = "N"
        End If
    End If
    
    If (Glo_Screen_No >= 2) Then
        If chk_Lane(1).value = 1 Then
            sChkLane2 = "Y"
        Else
            sChkLane2 = "N"
        End If
    End If
    
    If (Glo_Screen_No >= 4) Then
        If chk_Lane(2).value = 1 Then
            sChkLane3 = "Y"
        Else
            sChkLane3 = "N"
        End If
        If chk_Lane(3).value = 1 Then
            sChkLane4 = "Y"
        Else
            sChkLane4 = "N"
        End If
    End If
    
    If (Glo_Screen_No >= 6) Then
        If chk_Lane(4).value = 1 Then
            sChkLane5 = "Y"
        Else
            sChkLane5 = "N"
        End If
        If chk_Lane(5).value = 1 Then
            sChkLane6 = "Y"
        Else
            sChkLane6 = "N"
        End If
    End If
    
    If chk_Week(0).value = 1 Then
        sChkWeek1 = "Y"
    Else
        sChkWeek1 = "N"
    End If
    If chk_Week(1).value = 1 Then
        sChkWeek2 = "Y"
    Else
        sChkWeek2 = "N"
    End If
    If chk_Week(2).value = 1 Then
        sChkWeek3 = "Y"
    Else
        sChkWeek3 = "N"
    End If
    If chk_Week(3).value = 1 Then
        sChkWeek4 = "Y"
    Else
        sChkWeek4 = "N"
    End If
    If chk_Week(4).value = 1 Then
        sChkWeek5 = "Y"
    Else
        sChkWeek5 = "N"
    End If
    If chk_Week(5).value = 1 Then
        sChkWeek6 = "Y"
    Else
        sChkWeek6 = "N"
    End If
    If chk_Week(6).value = 1 Then
        sChkWeek7 = "Y"
    Else
        sChkWeek7 = "N"
    End If
    
    
    If (cmb_Rotation_YN.text = "����") Then
        sRotation = "Y"
    Else
        sRotation = "N"
    End If
    
    
    If (chk_App.value = 1) Then
        sApp = "Y"
        sApp_YesDate = Format(Now, "yyyy-mm-dd hh:nn:ss")
    Else
        sApp = "N"
        'sApp_YesDate = ""
    End If
    
    sApp_Certify_date = ""
    sLog_data = Format(Now, "yyyy-mm-dd hh:nn:ss")
    
    txt_CarNo.text = Replace(txt_CarNo.text, " ", "")
    tmpCarNo = txt_CarNo
    tmpName = txt_Name
    tmpPhone = txt_Phone
    tmpCarModel = txt_CarModel
    tmpObject = txt_Object
    tmpDong = Trim(cmb_Dong.text)
    tmpHo = Trim(txt_Ho)
    
    
    
    'stDate = Format(MaskEdBox_Start, "YYYYMMDD") & Format(DTPicker3, "hhnn") & "00"
    'edDate = Format(MaskEdBox_End, "YYYYMMDD") & Format(DTPicker4, "hhnn") & "59"
    stDate = Format(MaskEdBox_Start, "YYYYMMDD") & "000000"
    edDate = Format(MaskEdBox_End, "YYYYMMDD") & "235959"
    

    Call DBaseCheck
    
    If (CAR_NO_TMP = "") Then '�űԵ��

        If (sApp = "N") Then
            sQry = "INSERT INTO tb_reg (CAR_NO, CAR_MODEL, CAR_GUBUN, CAR_FEE, DRIVER_NAME, DRIVER_PHONE, DRIVER_DEPT, DRIVER_CLASS, START_DATE, END_DATE, ETC, REG_DATE, UPDATE_DATE, FEE_DATE,DAY_ROTATION_YN,REG_PART,LANE1,LANE2,LANE3,LANE4,LANE5,LANE6,WEEK1,WEEK2,WEEK3,WEEK4,WEEK5,WEEK6,WEEK7,ROTATION,APP_YN,APP_PW,APP_YES_DATE,APP_CERTIFY_DATE) "
            sQry = sQry & " VALUES ('" & tmpCarNo & "', '" & tmpCarModel & "', '" & cmb_Gubun.text & "', '" & MaskEdBox_Fee.text & "', '" & tmpName & "', '" & tmpPhone & "', '" & tmpDong & "', '" & tmpHo & "', '" & stDate & "', '" & edDate & "', '" & tmpObject & "', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "', '', '', '" & cmb_Rotation.text & "', '" & Glo_PartName & "', '" & sChkLane1 & "', '" & sChkLane2 & "', '" & sChkLane3 & "', '" & sChkLane4 & "', '" & sChkLane5 & "', '" & sChkLane6 & "', '" & sChkWeek1 & "', '" & sChkWeek2 & "', '" & sChkWeek3 & "', '" & sChkWeek4 & "', '" & sChkWeek5 & "', '" & sChkWeek6 & "', '" & sChkWeek7 & "', '" & sRotation & "', '" & sApp & "', '', Null, Null)"
        Else
            sQry = "INSERT INTO tb_reg (CAR_NO, CAR_MODEL, CAR_GUBUN, CAR_FEE, DRIVER_NAME, DRIVER_PHONE, DRIVER_DEPT, DRIVER_CLASS, START_DATE, END_DATE, ETC, REG_DATE, UPDATE_DATE, FEE_DATE,DAY_ROTATION_YN,REG_PART,LANE1,LANE2,LANE3,LANE4,LANE5,LANE6,WEEK1,WEEK2,WEEK3,WEEK4,WEEK5,WEEK6,WEEK7,ROTATION,APP_YN,APP_PW,APP_YES_DATE,APP_CERTIFY_DATE) "
            sQry = sQry & " VALUES ('" & tmpCarNo & "', '" & tmpCarModel & "', '" & cmb_Gubun.text & "', '" & MaskEdBox_Fee.text & "', '" & tmpName & "', '" & tmpPhone & "', '" & tmpDong & "', '" & tmpHo & "', '" & stDate & "', '" & edDate & "', '" & tmpObject & "', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "', '', '', '" & cmb_Rotation.text & "', '" & Glo_PartName & "', '" & sChkLane1 & "', '" & sChkLane2 & "', '" & sChkLane3 & "', '" & sChkLane4 & "', '" & sChkLane5 & "', '" & sChkLane6 & "', '" & sChkWeek1 & "', '" & sChkWeek2 & "', '" & sChkWeek3 & "', '" & sChkWeek4 & "', '" & sChkWeek5 & "', '" & sChkWeek6 & "', '" & sChkWeek7 & "', '" & sRotation & "', '" & sApp & "', '" & APP_INIT_PASSWORD & "', '" & sApp_YesDate & "', Null)"
        End If
        adoConn.Execute sQry
        
        sQry = "INSERT INTO tb_reg_log (CAR_NO, CAR_MODEL, CAR_GUBUN, CAR_FEE, DRIVER_NAME, DRIVER_PHONE, DRIVER_DEPT, DRIVER_CLASS, START_DATE, END_DATE, ETC, REG_DATE, UPDATE_DATE, FEE_DATE,DAY_ROTATION_YN,REG_PART,ACTION_LOG,AFTER_CAR_NO,ACTION_ID,  LANE1,LANE2,LANE3,LANE4,LANE5,LANE6,WEEK1,WEEK2,WEEK3,WEEK4,WEEK5,WEEK6,WEEK7,ROTATION,APP_YN,APP_PW,APP_YES_DATE,APP_CERTIFY_DATE,LOG_DATE) "
        sQry = sQry & " VALUES ('" & tmpCarNo & "', '" & tmpCarModel & "', '" & cmb_Gubun.text & "', '" & MaskEdBox_Fee.text & "', '" & tmpName & "', '" & tmpPhone & "', '" & tmpDong & "', '" & tmpHo & "', '" & stDate & "', '" & edDate & "', '" & tmpObject & "', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "', '', '', '" & cmb_Rotation.text & "', '" & Glo_PartName & "', '���', '', '" & Glo_Login_ID & "', '" & sChkLane1 & "', '" & sChkLane2 & "', '" & sChkLane3 & "', '" & sChkLane4 & "', '" & sChkLane5 & "', '" & sChkLane6 & "', '" & sChkWeek1 & "', '" & sChkWeek2 & "', '" & sChkWeek3 & "', '" & sChkWeek4 & "', '" & sChkWeek5 & "', '" & sChkWeek6 & "', '" & sChkWeek7 & "', '" & sRotation & "', '" & sApp & "', '" & APP_INIT_PASSWORD & "', '" & sApp_YesDate & "', '" & sApp_Certify_date & "', '" & sLog_data & "')"
        adoConn.Execute sQry
        
        
        List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_CarNo & "    ������� �Ϸ�", 0
        Call DataLogger("[REG Button]    " & txt_CarNo & "    �������� ��� �Ϸ�")
        
        
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' �����湮���� ��ɻ�� + ��/ȣ�� ���� => �����湮���� ���̺� ���
        If (Glo_GuestReg_YN = "Y") Then
            If (IsNumeric(tmpDong) = True And IsNumeric(tmpHo) = True) Then
                
                Dim sFormatDong As String
                Dim sFormatHo As String
                Dim sFormatID As String
                Dim maxparkcount As Integer '��������Ƚ��(��)
                Dim maxparktime As Integer '���������ð�(��)
                Dim maxparkday As Integer '�����������ϼ�
                Dim nowparkcount As Integer '1��~������� ����Ƚ��(��)
                Dim nowparktime As Integer '1��~������� �����ð�(��)
                
                
                maxparkcount = 1 '��������Ƚ��(��)
                maxparktime = 1 '���������ð�(��)
                maxparkday = 1 '�����������ϼ�
                nowparkcount = 0 '1��~������� ����Ƚ��(��)
                nowparktime = 0 '1��~������� �����ð�(��)
                
                sFormatDong = Format(Left(tmpDong, 3), "000")
                sFormatHo = Format(Left(tmpHo, 4), "0000")
                sFormatID = sFormatDong & sFormatHo
                
                If (isExist_GuestRegAdmin(sFormatID) = False) Then
                
                    '�⺻������ ��������(�����������ϼ�,��������Ƚ��,�����ð�)
                    Call GetParkPoint(maxparkday, maxparkcount, maxparktime)
                    
                    sQry = "INSERT INTO tb_guestreg_admin (VENDOR, SITE, NAME, ID, PASSWORD, CARNO, TEL, DRIVER_DEPT, DRIVER_CLASS, MAXPARKDAY,MAXPARKTIME,MAXPARKCOUNT,NOWPARKTIME,NOWPARKCOUNT, USE_YN, REG_DATE) "
                    sQry = sQry & " VALUES (0,0, '" & tmpName & "', '" & sFormatID & "', '0000', '" & tmpCarNo & "', '" & tmpPhone & "', '" & sFormatDong & "', '" & sFormatHo & "', " & maxparkday & " , " & maxparktime & " , " & maxparkcount & " , " & nowparktime & " , " & nowparkcount & " , 'Y', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "')"
                    adoConn.Execute sQry
                    
                    List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_CarNo & "    �����湮���� ������� �Ϸ�", 0
                    Call DataLogger("[REG Button]    " & txt_CarNo & "    �����湮���� �������� ��� �Ϸ�")
                Else
                    List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_CarNo & "   ���� �����湮���� �������", 0
                    Call DataLogger("[REG Button]    " & txt_CarNo & "    ���� �����湮���� �������")
                End If
            End If
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
        
        If (MaskEdBox_Fee <> "0") Then

            MBox.Label3.Caption = txt_CarNo.text & vbCrLf & MaskEdBox_Fee.text & "��"
            MBox.Label3.FontSize = 20
            MBox.Label1.Caption = "�� ������ ���������� ����մϴ�. ����Ͻðڽ��ϱ�?"
            MBox.Label2.Caption = "�������� ���� ���"
            MBox.Show 1
            If (Glo_MsgRet = True) Then
                adoConn.Execute "UPDATE tb_reg     SET FEE_DATE = '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "' WHERE CAR_NO = '" & txt_CarNo & "'"
            End If
        End If
    Else
        If (CAR_NO_TMP <> txt_CarNo) Then '���� ������ȣ ����

            If (sApp = "N") Then '������ȣ �� ����
                sQry = "UPDATE tb_reg SET CAR_NO='" & txt_CarNo & "',CAR_MODEL='" & tmpCarModel & "',CAR_GUBUN='" & cmb_Gubun & "', CAR_FEE='" & MaskEdBox_Fee.text & "',DRIVER_NAME='" & tmpName & "',DRIVER_PHONE='" & tmpPhone & "',DRIVER_DEPT='" & tmpDong & "',DRIVER_CLASS='" & tmpHo & "',START_DATE='" & stDate & "',END_DATE='" & edDate & "',ETC='" & tmpObject & "',UPDATE_DATE='" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "',DAY_ROTATION_YN='" & cmb_Rotation & "',LANE1='" & sChkLane1 & "',LANE2='" & sChkLane2 & "',LANE3='" & sChkLane3 & "',LANE4='" & sChkLane4 & "',LANE5='" & sChkLane5 & "',LANE6='" & sChkLane6 & "',WEEK1 ='" & sChkWeek1 & "',WEEK2='" & sChkWeek2 & "',WEEK3='" & sChkWeek3 & "',WEEK4='" & sChkWeek4 & "',WEEK5='" & sChkWeek5 & "',WEEK6='" & sChkWeek6 & "',WEEK7='" & sChkWeek7 & "',ROTATION='" & sRotation & "',APP_YN='" & sApp & "',APP_CERTIFY_DATE=Null WHERE CAR_NO='" & CAR_NO_TMP & "'"
            Else
                sQry = "UPDATE tb_reg SET CAR_NO ='" & txt_CarNo & "',CAR_MODEL='" & tmpCarModel & "',CAR_GUBUN='" & cmb_Gubun & "',CAR_FEE='" & MaskEdBox_Fee.text & "',DRIVER_NAME='" & tmpName & "',DRIVER_PHONE='" & tmpPhone & "',DRIVER_DEPT='" & tmpDong & "',DRIVER_CLASS='" & tmpHo & "',START_DATE='" & stDate & "',END_DATE='" & edDate & "',ETC='" & tmpObject & "',UPDATE_DATE='" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "',DAY_ROTATION_YN='" & cmb_Rotation & "',LANE1='" & sChkLane1 & "',LANE2='" & sChkLane2 & "',LANE3='" & sChkLane3 & "',LANE4='" & sChkLane4 & "',LANE5='" & sChkLane5 & "',LANE6='" & sChkLane6 & "',WEEK1 ='" & sChkWeek1 & "',WEEK2='" & sChkWeek2 & "',WEEK3='" & sChkWeek3 & "',WEEK4='" & sChkWeek4 & "',WEEK5='" & sChkWeek5 & "',WEEK6='" & sChkWeek6 & "',WEEK7='" & sChkWeek7 & "',ROTATION='" & sRotation & "',APP_YN='" & sApp & "',APP_YES_DATE='" & sApp_YesDate & "' WHERE CAR_NO='" & CAR_NO_TMP & "'"
            End If
            
            adoConn.Execute sQry
            
            sQry = "INSERT INTO tb_reg_log (CAR_NO, CAR_MODEL, CAR_GUBUN, CAR_FEE, DRIVER_NAME, DRIVER_PHONE, DRIVER_DEPT, DRIVER_CLASS, START_DATE, END_DATE, ETC, REG_DATE, UPDATE_DATE, FEE_DATE,DAY_ROTATION_YN,REG_PART,ACTION_LOG,AFTER_CAR_NO,ACTION_ID,  LANE1,LANE2,LANE3,LANE4,LANE5,LANE6,WEEK1,WEEK2,WEEK3,WEEK4,WEEK5,WEEK6,WEEK7,ROTATION,APP_YN,APP_PW,APP_YES_DATE,APP_CERTIFY_DATE,LOG_DATE) "
            sQry = sQry & " VALUES ('" & CAR_NO_TMP & "', '" & tmpCarModel & "', '" & cmb_Gubun.text & "', '" & MaskEdBox_Fee.text & "', '" & tmpName & "', '" & tmpPhone & "', '" & tmpDong & "', '" & tmpHo & "', '" & stDate & "', '" & edDate & "', '" & tmpObject & "', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "', '', '', '" & cmb_Rotation.text & "', '" & Glo_PartName & "', '����:������ȣ', '" & tmpCarNo & "', '" & Glo_Login_ID & "', '" & sChkLane1 & "', '" & sChkLane2 & "', '" & sChkLane3 & "', '" & sChkLane4 & "', '" & sChkLane5 & "', '" & sChkLane6 & "', '" & sChkWeek1 & "', '" & sChkWeek2 & "', '" & sChkWeek3 & "', '" & sChkWeek4 & "', '" & sChkWeek5 & "', '" & sChkWeek6 & "', '" & sChkWeek7 & "', '" & sRotation & "', '" & sApp & "', '" & APP_INIT_PASSWORD & "', '" & sApp_YesDate & "', '" & sApp_Certify_date & "', '" & sLog_data & "')"
            adoConn.Execute sQry
            
        Else
            If (sApp = "N") Then '������ȣ �� ����
                sQry = "UPDATE tb_reg     SET CAR_MODEL = '" & tmpCarModel & "', CAR_GUBUN = '" & cmb_Gubun & "', CAR_FEE = '" & MaskEdBox_Fee.text & "', DRIVER_NAME = '" & tmpName & "', DRIVER_PHONE = '" & tmpPhone & "', DRIVER_DEPT = '" & tmpDong & "', DRIVER_CLASS = '" & tmpHo & "', START_DATE = '" & stDate & "', END_DATE = '" & edDate & "', ETC = '" & tmpObject & "', UPDATE_DATE = '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "', DAY_ROTATION_YN = '" & cmb_Rotation & "', LANE1 = '" & sChkLane1 & "', LANE2= '" & sChkLane2 & "', LANE3= '" & sChkLane3 & "', LANE4= '" & sChkLane4 & "' , LANE5= '" & sChkLane5 & "', LANE6= '" & sChkLane6 & "', WEEK1 = '" & sChkWeek1 & "', WEEK2 = '" & sChkWeek2 & "' , WEEK3 = '" & sChkWeek3 & "' , WEEK4 = '" & sChkWeek4 & "' , WEEK5 = '" & sChkWeek5 & "' , WEEK6 = '" & sChkWeek6 & "' , WEEK7 = '" & sChkWeek7 & "', ROTATION = '" & sRotation & "', APP_YN='" & sApp & "', APP_CERTIFY_DATE=Null WHERE CAR_NO='" & CAR_NO_TMP & "'"
            Else
                sQry = "UPDATE tb_reg     SET CAR_MODEL = '" & tmpCarModel & "', CAR_GUBUN = '" & cmb_Gubun & "', CAR_FEE = '" & MaskEdBox_Fee.text & "', DRIVER_NAME = '" & tmpName & "', DRIVER_PHONE = '" & tmpPhone & "', DRIVER_DEPT = '" & tmpDong & "', DRIVER_CLASS = '" & tmpHo & "', START_DATE = '" & stDate & "', END_DATE = '" & edDate & "', ETC = '" & tmpObject & "', UPDATE_DATE = '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "', DAY_ROTATION_YN = '" & cmb_Rotation & "', LANE1 = '" & sChkLane1 & "', LANE2= '" & sChkLane2 & "', LANE3= '" & sChkLane3 & "', LANE4= '" & sChkLane4 & "' , LANE5= '" & sChkLane5 & "', LANE6= '" & sChkLane6 & "', WEEK1 = '" & sChkWeek1 & "', WEEK2 = '" & sChkWeek2 & "' , WEEK3 = '" & sChkWeek3 & "' , WEEK4 = '" & sChkWeek4 & "' , WEEK5 = '" & sChkWeek5 & "' , WEEK6 = '" & sChkWeek6 & "' , WEEK7 = '" & sChkWeek7 & "', ROTATION = '" & sRotation & "', APP_YN='" & sApp & "', APP_YES_DATE='" & sApp_YesDate & "' WHERE CAR_NO='" & CAR_NO_TMP & "'"
            End If
            adoConn.Execute sQry
            
            sQry = "INSERT INTO tb_reg_log (CAR_NO, CAR_MODEL, CAR_GUBUN, CAR_FEE, DRIVER_NAME, DRIVER_PHONE, DRIVER_DEPT, DRIVER_CLASS, START_DATE, END_DATE, ETC, REG_DATE, UPDATE_DATE, FEE_DATE,DAY_ROTATION_YN,REG_PART,ACTION_LOG,AFTER_CAR_NO,ACTION_ID,  LANE1,LANE2,LANE3,LANE4,LANE5,LANE6,WEEK1,WEEK2,WEEK3,WEEK4,WEEK5,WEEK6,WEEK7,ROTATION,APP_YN,APP_PW,APP_YES_DATE,APP_CERTIFY_DATE,LOG_DATE) "
            sQry = sQry & " VALUES ('" & txt_CarNo & "', '" & tmpCarModel & "', '" & cmb_Gubun.text & "', '" & MaskEdBox_Fee.text & "', '" & tmpName & "', '" & tmpPhone & "', '" & tmpDong & "', '" & tmpHo & "', '" & stDate & "', '" & edDate & "', '" & tmpObject & "', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "', '', '', '" & cmb_Rotation.text & "', '" & Glo_PartName & "', '����:������ȣ ��', '', '" & Glo_Login_ID & "', '" & sChkLane1 & "', '" & sChkLane2 & "', '" & sChkLane3 & "', '" & sChkLane4 & "', '" & sChkLane5 & "', '" & sChkLane6 & "', '" & sChkWeek1 & "', '" & sChkWeek2 & "', '" & sChkWeek3 & "', '" & sChkWeek4 & "', '" & sChkWeek5 & "', '" & sChkWeek6 & "', '" & sChkWeek7 & "', '" & sRotation & "', '" & sApp & "', '" & APP_INIT_PASSWORD & "', '" & sApp_YesDate & "', '" & sApp_Certify_date & "', '" & sLog_data & "')"
            adoConn.Execute sQry
            
        End If
        List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_CarNo & "    �������� ���� �Ϸ�", 0
        Call DataLogger("[REG Button]    " & txt_CarNo & "    �������� ���� �Ϸ�")
        
        
        

    End If
    
    cmd_PWInit.Enabled = False
    
    RegQry = "SELECT * From tb_reg ORDER BY CAR_NO"
    Call ListView_REG_Draw
    Call ListView_REG_SQL
    
    Exit Sub

Err_p:
    List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & Err.Description, 0
    Call DataLogger("[REG Insert_Record]    " & Err.Description)
'On Error Resume Next
'    If (Err = 3022) Then
'        Msg_Box.Label2.Caption = "������ ���̽� ����"
'        Msg_Box.Label1.Caption = "�ߺ��� ������ȣ�� ��������ʽ��ϴ�."
'        Msg_Box.Show 1
'    End If

End Sub


'�����湮���� ������ ���̺��� �ش� ���̵� ���� �������� üũ
Private Function isExist_GuestRegAdmin(sID As String)
    Dim rs As Recordset
    Dim qry As String
    
    qry = "SELECT * from tb_guestreg_admin WHERE ID = '" & sID & "'"
    Set rs = New ADODB.Recordset
    rs.Open qry, adoConn

    If rs.EOF Then
        isExist_GuestRegAdmin = False
    Else
        isExist_GuestRegAdmin = True
    End If
    
    Set rs = Nothing
    
End Function

Private Function GetParkPoint(maxparkday As Integer, maxparkcount As Integer, maxparktime As Integer)
    Dim rs As Recordset
    Dim qry As String
    
    qry = "SELECT * FROM tb_config WHERE NAME = 'GuestCarReg_MaxParkCount' OR NAME = 'GuestCarReg_MaxParkDay' OR NAME = 'GuestCarReg_MaxParkTime'"
    Set rs = New ADODB.Recordset
    rs.Open qry, adoConn

    Do While (Not rs.EOF)
        If (rs!name = "GuestCarReg_MaxParkDay") Then
            maxparkday = rs!Content
        ElseIf (rs!name = "GuestCarReg_MaxParkCount") Then
            maxparkcount = rs!Content
        ElseIf (rs!name = "GuestCarReg_MaxParkTime") Then
            maxparktime = rs!Content
        End If
        rs.MoveNext
    Loop

    Set rs = Nothing
    
End Function


Private Sub cmd_Button_Click(Index As Integer)
    Dim i, j As Integer
    Dim myExcelFile As New ExcelFile
    Dim tmpFileName As String

    Dim rs As Recordset
    Dim qry As String
    Dim sQry As String
    Dim bQryResult As Boolean
    Dim tmpCarNo, tmpName, tmpCarModel, tmpObject, tmpDong, tmpHo As String
    
    tmpCarNo = lbl_Carno_tmp
    tmpName = lbl_Name_tmp
    tmpDong = lbl_dong_tmp
    tmpHo = lbl_ho_tmp
    tmpObject = lbl_object_tmp
    tmpCarModel = lbl_CarModel_tmp

    Select Case Index
        Case 0  '����
            List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    �������/���� ����", 0
            Call DataLogger("[REG Button]    " & txt_CarNo & "    �������/���� ����")
            'Call SaveReg2
            Unload Me
            'Me.Hide
            Exit Sub
           
        Case 1  '�ű��Է�
            If (CAR_NO_TMP = "") Then
                If (Data_Error_Check = False) Then
                    Msg_Box.Label2.Caption = "�ʵ� �Է� ����"
                    Msg_Box.Label1.Caption = "�߿� �׸��� ��Ȯ�� �Է��ϼ���."
                    Msg_Box.Show 1
                    Call Clear_Field
                Else
                    Call Insert_Record
                    Call Clear_Field
                    Call SaveReg2
                End If
            Else
                Msg_Box.Label2.Caption = "�ű� ������ �Է� ����"
                Msg_Box.Label1.Caption = "�ű� �����Ͱ� �ƴմϴ�." & vbCrLf & vbCrLf & " �ٽ� �ѹ� Ȯ���ϼ���."
                Msg_Box.Show 1
                Call Clear_Field
            End If
            Exit Sub
        
        Case 2  '����
            If (CAR_NO_TMP = "") Then
               Call Clear_Field
               Exit Sub
            End If
            If (CAR_NO_TMP <> Me.txt_CarNo) Then
                Msg_Box.Label2.Caption = "������ ���� ����"
                Msg_Box.Label1.Caption = "������ �����͸� �ٽ� ������ �ֽʽÿ�."
                Msg_Box.Show 1
                Exit Sub
            End If
            MBox.Label3.Caption = txt_CarNo.text
            MBox.Label1.Caption = "�� ������ ������� ������ �����մϴ�." & vbCrLf & vbCrLf & " �����Ͻðڽ��ϱ�?"
            MBox.Label2.Caption = "������� ���� ����"
            MBox.Show 1
            If (Glo_MsgRet = True) Then
               Call Delete_Record
               Call SaveReg2
            End If
            Call Clear_Field
            Exit Sub
            
        Case 3   '�ʱ�ȭ
            Call Clear_Field
            Exit Sub
                
        Case 4  '����
            If (CAR_NO_TMP = "") Then
                Msg_Box.Label2.Caption = "�ʵ� ����"
                Msg_Box.Label1.Caption = "�ű� ����ڷ� �Դϴ�." & vbCrLf & vbCrLf & " �ٽ� Ȯ�� �ϼ���."
                Msg_Box.Show 1
                Exit Sub
            Else
                If (txt_CarNo.text = CAR_NO_TMP) Then
                    If (Data_Error_Check = False) Then
                        Msg_Box.Label2.Caption = "�ʵ� �Է� ����"
                        Msg_Box.Label1.Caption = "�߿��� �׸��� ���� �Ǵ� �߸� �Է��Ͽ����ϴ�."
                        Msg_Box.Show 1
                    Else
                        MBox.Label3.Caption = txt_CarNo.text
                        MBox.Label1.Caption = "�����Ͻ� ������� ������ ����˴ϴ�." & vbCrLf & vbCrLf & " ���� �Ͻðڽ��ϱ�?"
                        MBox.Label2.Caption = "������� �ڷ� ����"
                        MBox.Show 1
                        If (Glo_MsgRet = True) Then
                           Call Insert_Record
                           Call Clear_Field
                           Call SaveReg2
                           'txt_CarNo.SetFocus
                        End If
                    End If
                Else
                    If (Data_Error_Check = False) Then
                        Msg_Box.Label2.Caption = "�ʵ� �Է� ����"
                        Msg_Box.Label1.Caption = "�߿��� �׸��� ���� �Ǵ� �߸� �Է��Ͽ����ϴ�."
                        Msg_Box.Show 1
                    Else
                        MBox.Label3.Caption = tmpCarNo
                        MBox.Label1.Caption = "�����Ͻ� �ڷ��� ������ȣ�� ����˴ϴ�." & vbCrLf & vbCrLf & " ���� �Ͻðڽ��ϱ�?"
                        MBox.Label2.Caption = "������� ���� ����"
                        MBox.Show 1
                        If (Glo_MsgRet = True) Then
                           Call Insert_Record
                           Call Clear_Field
                           Call SaveReg2
                           'txt_CarNo.SetFocus
                        End If
                    End If
                End If
            End If
            Exit Sub
    
        Case 5
            Call SaveReg
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
                        
                        sQry = "UPDATE tb_reg SET FEE_DATE = '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "' WHERE CAR_NO = '" & txt_CarNo & "'"
                        bQryResult = DataBaseQueryExec(adoConn, sQry, NWERR_GATE_STAY)
                        If (bQryResult = False) Then
                            List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    ��Ʈ��ũ �� DB ���˹ٶ��ϴ�", 0
                            Call DataLogger("[FrmReg cmd_Button7]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
                            Exit Sub
                        End If
                        
                        sQry = "INSERT INTO TB_FEE VALUES ('" & tmpCarNo & "', '" & tmpCarModel & "', '" & cmb_Gubun & "', '" & MaskEdBox_Fee.text & "', '" & tmpName & "', '" & txt_Phone & "', '" & cmb_Dong & "', '" & tmpHo & "', '" & Format(MaskEdBox_Start, "YYYYMMDD") & "', '" & Format(MaskEdBox_End, "YYYYMMDD") & "', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "')"
                        bQryResult = DataBaseQueryExec(adoConn, sQry, NWERR_GATE_STAY)
                        If (bQryResult = False) Then
                            List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    ��Ʈ��ũ �� DB ���˹ٶ��ϴ�", 0
                            Call DataLogger("[FrmReg cmd_Button7]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
                            Exit Sub
                        End If
                
                
                
                        List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & tmpCarNo & "    " & MaskEdBox_Fee.text & "��    �������� �Ϸ�", 0
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
            Call ListView_REG_Draw
            Call ListView_REG_SQL
            Exit Sub
            
        Case 8  '������̷�
            'Unload Me
            'FrmRegHistory.Show 1
            FrmRegHistory.Show 0
            Me.MousePointer = 0
            Call DataLogger("[HOST Button]    " & "����� �̷� ȭ�� ����")
            Exit Sub
    End Select

On Error Resume Next

End Sub


'�ʼ� �Է� ������ Ȯ��
Private Function Data_Error_Check()
    Dim Error_Flag As Boolean
    Dim i As Integer
    Dim iChkLane As Integer
        
    Error_Flag = True
    
    
    txt_CarNo.text = Replace(txt_CarNo.text, " ", "")
'    txt_CarNo.Text = Trim(txt_CarNo.Text)
    
    If Not ((LenH(txt_CarNo.text) = 11) Or (LenH(txt_CarNo.text) = 12) Or (LenH(txt_CarNo.text) = 8) Or (LenH(txt_CarNo.text) = 9)) Then
        Error_Flag = False
    End If
    If (LenH(txt_CarNo.text) = 0) Then
        Error_Flag = False
    End If
    If (IsNumeric(MaskEdBox_Fee.text) = False) Then
        MaskEdBox_Fee.text = "0"
        'Error_Flag = False
    End If
    
    If (Glo_User_Type = "����1/����2") Then
        If (LenH(txt_Ho.text) = 0) Then
            'txt_Phone.Text = " "
            'Error_Flag = False
        Else
            txt_Ho.text = Mid(txt_Ho.text, 1, 16)
        End If
        If (LenH(cmb_Dong.text) = 0) Then
            'txt_CarModel.Text = " "
            'Error_Flag = False
        Else
            cmb_Dong.text = MidH(cmb_Dong.text, 1, 16)
        End If
    Else
    End If
    
    If (IsDate(MaskEdBox_Start.text) = False) Then
        Error_Flag = False
    End If
    If (IsDate(MaskEdBox_End.text) = False) Then
        Error_Flag = False
    End If
    If (Len(txt_Object.text) = 0) Then
        txt_Object.text = " "
        'Error_Flag = False
    Else
        txt_Object.text = MidH(txt_Object.text, 1, 64)
    End If
    
    iChkLane = 0
    For i = 0 To 5
        If (chk_Lane(i).value = 1) Then
            iChkLane = iChkLane + 1
        End If
    Next i
    If iChkLane = 0 Then
        Error_Flag = False
    End If
    
    
    Data_Error_Check = Error_Flag

End Function


'�����湮����
Private Sub SSCommand1_Click()
        
End Sub

Private Sub txt_CarNo_Change()

    'If (LenH(txt_CarNo) > 7 Or LenH(txt_CarNo) = 4) Then
        Call Search_Record
    'End If
End Sub

Sub Search_Record()
    Dim rs As Recordset
    Dim SQL_SEARCH As String
    Dim itmX As ListItem
    Dim INDEX_NO As Long
    
    Dim bQryResult As Boolean

On Error GoTo Err_p

    If (lbl_gubun_tmp = "�湮����") Then
        SQL_SEARCH = "SELECT * From tb_guestReg WHERE CAR_NO = '" & txt_CarNo & "' ORDER BY CAR_NO"
    Else
        SQL_SEARCH = "SELECT * From tb_reg WHERE CAR_NO = '" & txt_CarNo & "' ORDER BY CAR_GUBUN"
    End If
    
    
    Set rs = New ADODB.Recordset
    'rs.Open SQL_SEARCH, adoConn
    
     bQryResult = DataBaseQuery(rs, adoConn, SQL_SEARCH, False)
     If (bQryResult = False) Then
        List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    ��Ʈ��ũ �� DB ���˹ٶ��ϴ�", 0
        Call DataLogger("[FrmReg Search_Record]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
        Exit Sub
    End If
    

    
    If (rs.RecordCount <> 0) Then
'        cmd_Button(3).Enabled = True
'        cmd_Button(1).Enabled = True
'        cmd_Button(4).Enabled = True
'        cmd_Button(2).Enabled = True
    
    
        CAR_NO_TMP = rs!CAR_NO
        'INDEX_NO_TMP = ListView_REG.SelectedItem.Text
        'txt_Num = "" & rs!CAR_NO
        txt_Num = "" & Format(rs!REG_DATE, "yyyy-mm-dd hh:nn:ss")
        txt_Name = "" & rs!DRIVER_NAME
        txt_Phone = "" & rs!DRIVER_PHONE
        txt_CarModel = "" & rs!CAR_MODEL
        MaskEdBox_Fee.text = "" & rs!CAR_FEE
        'txt_Dong = "" & rs!DRIVER_DEPT
        cmb_Dong = "" & rs!DRIVER_DEPT
        txt_Ho = "" & rs!DRIVER_CLASS
        
        'MaskEdBox_Start.text = Format(Left(rs!Start_Date, 8), "####-##-##")
        'MaskEdBox_End.text = Format(Left(rs!End_Date, 8), "####-##-##")
        MaskEdBox_Start.text = Left(rs!START_DATE, 10)
        MaskEdBox_End.text = Left(rs!END_DATE, 10)
        
'        If (Len(rs!Start_Date) = 8) Then
'            DTPicker3.value = "00:00:00"
'        Else
'            DTPicker3.value = Format(Mid(rs!Start_Date, 9, 4), "00:00")
'        End If
'        If (Len(rs!End_Date) = 8) Then
'            DTPicker4.value = "23:59:59"
'        Else
'            DTPicker4.value = Format(Mid(rs!End_Date, 9, 4), "00:00")
'        End If

        DTPicker3.value = "00:00:00"
        DTPicker4.value = "23:59:59"
        
        
        Select Case rs!DAY_ROTATION_YN
            Case "����"
                cmb_Rotation.ListIndex = 0
            Case Else
                cmb_Rotation.ListIndex = 1
        End Select
        txt_Object = "" & rs!ETC
        
        If (rs!LANE1 = "Y") Then
            chk_Lane(0).value = 1
        Else
            chk_Lane(0).value = 0
        End If
        
        If (rs!LANE2 = "Y") Then
            chk_Lane(1).value = 1
        Else
            chk_Lane(1).value = 0
        End If
        
        If (rs!LANE3 = "Y") Then
            chk_Lane(2).value = 1
        Else
            chk_Lane(2).value = 0
        End If
        
        If (rs!LANE4 = "Y") Then
            chk_Lane(3).value = 1
        Else
            chk_Lane(3).value = 0
        End If
        
        If (rs!LANE5 = "Y") Then
            chk_Lane(4).value = 1
        Else
            chk_Lane(4).value = 0
        End If
        
        If (rs!LANE6 = "Y") Then
            chk_Lane(5).value = 1
        Else
            chk_Lane(5).value = 0
        End If
        
        
        If (rs!WEEK1 = "Y") Then
            chk_Week(0).value = 1
        Else
            chk_Week(0).value = 0
        End If
        If (rs!WEEK2 = "Y") Then
            chk_Week(1).value = 1
        Else
            chk_Week(1).value = 0
        End If
        If (rs!WEEK3 = "Y") Then
            chk_Week(2).value = 1
        Else
            chk_Week(2).value = 0
        End If
        If (rs!WEEK4 = "Y") Then
            chk_Week(3).value = 1
        Else
            chk_Week(3).value = 0
        End If
        If (rs!WEEK5 = "Y") Then
            chk_Week(4).value = 1
        Else
            chk_Week(4).value = 0
        End If
        If (rs!WEEK6 = "Y") Then
            chk_Week(5).value = 1
        Else
            chk_Week(5).value = 0
        End If
        If (rs!WEEK7 = "Y") Then
            chk_Week(6).value = 1
        Else
            chk_Week(6).value = 0
        End If
        
        
        If (rs!Rotation = "Y") Then
            cmb_Rotation_YN.ListIndex = 0
        Else
            cmb_Rotation_YN.ListIndex = 1
        End If
        
        
        If (rs!APP_YN = "Y") Then
            chk_App.value = 1
        Else
            chk_App.value = 0
        End If
        
        
        
        cmb_Gubun.text = "" & rs!CAR_GUBUN
        
        
        
    Else
        'Call Clear_Field
    End If
    Set rs = Nothing
    
    
    
    Exit Sub
    
Err_p:
    List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    [Search Record]  " & Err.Description, 0
    Set rs = Nothing
    Call DataLogger(" [Search Record]  " & Err.Description)
End Sub


Private Sub cmd_Search_Click()

'''    If Option1(0).value = True Then
'''        If Len(txt_tmpCarNo) <> 0 Then
'''            Select Case cmb_GB.ListIndex
'''                Case 0
'''                    RegQry = "SELECT * From tb_reg Where CAR_NO Like '%" & txt_tmpCarNo & "'"
'''                Case 1
'''                    RegQry = "SELECT * From tb_reg Where DRIVER_NAME Like '%" & txt_tmpCarNo & "%'"
'''                Case 2
'''                    RegQry = "SELECT * From tb_reg Where DRIVER_DEPT Like '%" & txt_tmpCarNo & "%'"
'''                Case 3
'''                    RegQry = "SELECT * From tb_reg Where DRIVER_CLASS Like '%" & txt_tmpCarNo & "%'"
'''                Case Else
'''                    RegQry = "SELECT * From tb_reg Where CAR_GUBUN Like '%" & txt_tmpCarNo & "%'"
'''            End Select
'''        Else
'''            Select Case cmb_GB.ListIndex
'''                Case 0
'''                    RegQry = "SELECT * From tb_reg Order By CAR_NO"
'''                Case 1
'''                    RegQry = "SELECT * From tb_reg Order By DRIVER_NAME"
'''                Case 2
'''                    RegQry = "SELECT * From tb_reg Order By DRIVER_DEPT"
'''                Case 3
'''                    RegQry = "SELECT * From tb_reg Order By DRIVER_CLASS"
'''                Case Else
'''                    RegQry = "SELECT * From tb_reg Order By CAR_GUBUN"
'''            End Select
'''        End If
'''    Else
'''        If Len(cmbDong.Text) = 0 Then
'''            If Len(cmbHo.Text) = 0 Then
'''                RegQry = "SELECT * From tb_reg"
'''            Else
'''                RegQry = "SELECT * From tb_reg Where DRIVER_CLASS = '" & cmbHo.Text & "'"
'''            End If
'''        Else
'''            If Len(cmbHo.Text) = 0 Then
'''                RegQry = "SELECT * From tb_reg Where DRIVER_DEPT = '" & cmbDong.Text & "'"
'''            Else
'''                RegQry = "SELECT * From tb_reg Where DRIVER_DEPT = '" & cmbDong.Text & "' AND DRIVER_CLASS = '" & cmbHo.Text & "'"
'''            End If
'''        End If
'''    End If
'''
'''    txt_tmpCarNo = ""
'''    Call Clear_Field
'''    Call ListView_REG_Draw
'''    Call ListView_REG_SQL

    If (cmb_GB.text = "�Ҽ�/����" Or cmb_GB.text = "��/ȣ��") Then
        If Len(cmbDong.text) = 0 Then
            If Len(cmbHo.text) = 0 Then
                RegQry = "SELECT * From tb_reg ORDER BY DRIVER_DEPT, DRIVER_CLASS "
            Else
                RegQry = "SELECT * From tb_reg Where DRIVER_CLASS = '" & cmbHo.text & "' ORDER BY DRIVER_DEPT, DRIVER_CLASS "
            End If
        Else
            If Len(cmbHo.text) = 0 Then
                RegQry = "SELECT * From tb_reg Where DRIVER_DEPT = '" & cmbDong.text & "' ORDER BY DRIVER_DEPT, DRIVER_CLASS "
            Else
                RegQry = "SELECT * From tb_reg Where DRIVER_DEPT = '" & cmbDong.text & "' AND DRIVER_CLASS = '" & cmbHo.text & "' ORDER BY DRIVER_DEPT, DRIVER_CLASS "
            End If
        End If
    
    ElseIf (cmb_GB.text = "�� ��") Then
        If (cmb_GubunSrch = "") Then
                RegQry = "SELECT * From tb_reg ORDER BY CAR_GUBUN "
        Else
                RegQry = "SELECT * From tb_reg Where CAR_GUBUN Like '%" & cmb_GubunSrch.text & "%' ORDER BY CAR_GUBUN"
        End If
    
    ElseIf (cmb_GB.text = "�湮����") Then
        If Len(cmbDong.text) = 0 Then
            If Len(cmbHo.text) = 0 Then
                'RegQry = "SELECT * From tb_reg Where CAR_GUBUN ='�湮����' ORDER BY DRIVER_DEPT, DRIVER_CLASS "
                RegQry = "SELECT * From tb_guestReg Where CAR_GUBUN ='�湮����' ORDER BY DRIVER_DEPT, DRIVER_CLASS "
            Else
                'RegQry = "SELECT * From tb_reg Where CAR_GUBUN ='�湮����' AND DRIVER_CLASS = '" & cmbHo.text & "' ORDER BY DRIVER_DEPT, DRIVER_CLASS "
                RegQry = "SELECT * From tb_guestReg Where CAR_GUBUN ='�湮����' AND DRIVER_CLASS = '" & cmbHo.text & "' ORDER BY DRIVER_DEPT, DRIVER_CLASS "
            End If
        Else
            If Len(cmbHo.text) = 0 Then
                'RegQry = "SELECT * From tb_reg Where CAR_GUBUN ='�湮����' AND DRIVER_DEPT = '" & cmbDong.text & "' ORDER BY DRIVER_DEPT, DRIVER_CLASS "
                RegQry = "SELECT * From tb_guestReg Where CAR_GUBUN ='�湮����' AND DRIVER_DEPT = '" & cmbDong.text & "' ORDER BY DRIVER_DEPT, DRIVER_CLASS "
            Else
                'RegQry = "SELECT * From tb_reg Where CAR_GUBUN ='�湮����' AND DRIVER_DEPT = '" & cmbDong.text & "' AND DRIVER_CLASS = '" & cmbHo.text & "' ORDER BY DRIVER_DEPT, DRIVER_CLASS "
                RegQry = "SELECT * From tb_guestReg Where CAR_GUBUN ='�湮����' AND DRIVER_DEPT = '" & cmbDong.text & "' AND DRIVER_CLASS = '" & cmbHo.text & "' ORDER BY DRIVER_DEPT, DRIVER_CLASS "
            End If
        End If
    
    Else
        
        Select Case cmb_GB.text
            Case ""
                RegQry = "SELECT * From tb_reg "
            Case "������ȣ"
                RegQry = "SELECT * From tb_reg Where CAR_NO Like '%" & txt_tmpCarNo & "%' ORDER BY CAR_NO"
            Case "�� ��"
                RegQry = "SELECT * From tb_reg Where DRIVER_NAME Like '%" & txt_tmpCarNo & "%' ORDER BY DRIVER_NAME"
        End Select

    End If
    
    txt_tmpCarNo = ""
    Call Clear_Field
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
    
    On Error Resume Next
    
    If (KeyAscii = vbKeySpace) Then
        If FrmReg.ActiveControl = txt_CarNo Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
    
    If (KeyAscii = 44) Then ' ,(��ǥ,44)�� �Էµ� �����͸� csv ���·� �����ϸ�, csv������ �ٽ� �ε��Ҷ� �����߻� ���ɼ�����.
            KeyAscii = 0
            Exit Sub
    End If

    If (KeyAscii = 13) Then
        If ((Len(txt_tmpCarNo) <> 0) Or (Len(cmb_GubunSrch) <> 0)) Then
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

Private Sub SaveReg()
Dim tmpFileName As String
On Error GoTo Err_p
    tmpFileName = Format(Now, "YYYYMMDD_HHMMSS")
    tmpFileName = App.Path & "\Excel\" & tmpFileName & "_�������_" & cmb_Search.text
        
        
    CommonDialog1.CancelError = True
    CommonDialog1.InitDir = App.Path
    CommonDialog1.Filter = "��������(*.csv)|*.csv"
    CommonDialog1.fileName = tmpFileName
    CommonDialog1.ShowSave
    tmpFileName = CommonDialog1.fileName
    tmpFileName = Mid(tmpFileName, 1, Len(tmpFileName) - 4)

    Call MakeCSV(ListView_REG, tmpFileName)
    Exit Sub
Err_p:
     Select Case Err
    Case 32755 '  Dialog Cancelled
        'MsgBox "you cancelled the dialog box"
    Case Else
        'MsgBox "Unexpected error. Err " & Err & " : " & Error
    End Select
End Sub

Private Sub SaveReg2()
    Dim tmpFileName As String
    Dim sCmd As String

On Error GoTo Err_p
    tmpFileName = Format(Now, "YYYYMMDD")
    tmpFileName = App.Path & "\Backup\" & tmpFileName & "_�������"

    If (IsFile(tmpFileName & ".CSV") = True) Then
        Kill tmpFileName & ".CSV"
    End If
    
    Call MakeCSV(ListView_REG, tmpFileName)
    Exit Sub
Err_p:
     Select Case Err
    Case 32755 '  Dialog Cancelled
        'MsgBox "you cancelled the dialog box"
    Case Else
        'MsgBox "Unexpected error. Err " & Err & " : " & Error
    End Select
End Sub

Public Sub CtrlEnable(ByVal sContens As String, ByVal bEnable As Boolean)
    
End Sub

Private Sub Set_cmbDong()
    Dim bQryResult As Boolean
    Dim rs As Recordset
    Dim qry As String
On Error GoTo Err_p

    qry = "SELECT tb_reg.DRIVER_DEPT From tb_reg Group By tb_reg.DRIVER_DEPT"

    Set rs = New ADODB.Recordset
'    rs.Open Qry, adoConn
     bQryResult = DataBaseQuery(rs, adoConn, qry, False)
     If (bQryResult = False) Then
        List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    ��Ʈ��ũ �� DB ���˹ٶ��ϴ�", 0
        Call DataLogger("[FrmReg]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
        Exit Sub
    End If
    
    cmbDong.Clear
    cmb_Dong.Clear
    If Not rs.EOF Then
        Do While Not (rs.EOF)
            cmbDong.AddItem "" & rs!DRIVER_DEPT
            cmb_Dong.AddItem "" & rs!DRIVER_DEPT
            
            'List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & rs!DRIVER_Dept, 0
            
            rs.MoveNext
        Loop
    End If
    Set rs = Nothing

Exit Sub
Err_p:
    Call DataLogger("[FrmReg Set_cmbDong]    " & Err.Description)
End Sub

Private Sub Set_cmbHo()
    Dim bQryResult As Boolean
    Dim rs As Recordset
    Dim qry As String
On Error GoTo Err_p

    qry = "SELECT tb_reg.DRIVER_CLASS From tb_reg Group By tb_reg.DRIVER_CLASS"
    
    Set rs = New ADODB.Recordset
'    rs.Open Qry, adoConn
     bQryResult = DataBaseQuery(rs, adoConn, qry, False)
     If (bQryResult = False) Then
        List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    ��Ʈ��ũ �� DB ���˹ٶ��ϴ�", 0
        Call DataLogger("[FrmReg]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
        Exit Sub
    End If
    
    cmbHo.Clear
    If Not rs.EOF Then
        Do While Not (rs.EOF)
            cmbHo.AddItem "" & rs!DRIVER_CLASS
            rs.MoveNext
        Loop
    End If
    Set rs = Nothing
Exit Sub

Err_p:
    Call DataLogger("[FrmReg Set_cmbHo]    " & Err.Description)
End Sub

Private Sub Set_cmbGubunSrch()
    Dim sQry As String
    Dim bQryResult As String
On Error GoTo Err_p
    sQry = "SELECT tb_reg.CAR_GUBUN From tb_reg Group By tb_reg.CAR_GUBUN"

    Set rs = New ADODB.Recordset
     bQryResult = DataBaseQuery(rs, adoConn, sQry, False)
     If (bQryResult = False) Then
        List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    ��Ʈ��ũ �� DB ���˹ٶ��ϴ�", 0
        Call DataLogger("[FrmReg]    " & "��Ʈ��ũ �� DB ���˹ٶ��ϴ�")
        Exit Sub
    End If
    
    cmb_GubunSrch.Clear
    If Not rs.EOF Then
        Do While Not (rs.EOF)
            cmb_GubunSrch.AddItem "" & rs!CAR_GUBUN
            rs.MoveNext
        Loop
    End If
    Set rs = Nothing
Exit Sub

Err_p:
    Call DataLogger("[FrmReg Set_cmbGubunSrch]    " & Err.Description)
End Sub



