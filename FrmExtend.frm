VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmExtend 
   BackColor       =   &H00404040&
   BorderStyle     =   1  '���� ����
   Caption         =   "ParkingManager��"
   ClientHeight    =   11010
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   14160
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11010
   ScaleWidth      =   14160
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.Frame Frame16 
      BackColor       =   &H00404040&
      Caption         =   " 17. ����Ͼ˸� "
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   9510
      TabIndex        =   91
      Top             =   2730
      Visible         =   0   'False
      Width           =   4470
      Begin VB.CheckBox chk_MobileAlarm 
         BackColor       =   &H00404040&
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   270
         TabIndex        =   92
         Top             =   300
         Width           =   1080
      End
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00404040&
      Caption         =   " 16. �����湮���� "
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   9510
      TabIndex        =   85
      Top             =   1740
      Width           =   4470
      Begin VB.CheckBox chk_GuestCarReg 
         BackColor       =   &H00404040&
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   270
         TabIndex        =   86
         Top             =   300
         Width           =   1080
      End
   End
   Begin VB.Frame Frame14 
      BackColor       =   &H00404040&
      Caption         =   " 15. ������ "
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   9510
      TabIndex        =   83
      Top             =   870
      Width           =   4470
      Begin VB.CheckBox chk_WebDC 
         BackColor       =   &H00404040&
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   240
         TabIndex        =   84
         Top             =   300
         Width           =   1110
      End
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H00404040&
      Caption         =   " 3. ���ܱ� �ڵ����� [�ڸ����] "
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1125
      Left            =   180
      TabIndex        =   76
      Top             =   3150
      Width           =   4470
      Begin VB.CheckBox chk_NoWork_YN 
         BackColor       =   &H00404040&
         Caption         =   "����6"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   5
         Left            =   2970
         TabIndex        =   82
         Top             =   720
         Width           =   1350
      End
      Begin VB.CheckBox chk_NoWork_YN 
         BackColor       =   &H00404040&
         Caption         =   "����5"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   4
         Left            =   1560
         TabIndex        =   81
         Top             =   720
         Width           =   1350
      End
      Begin VB.CheckBox chk_NoWork_YN 
         BackColor       =   &H00404040&
         Caption         =   "����4"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   3
         Left            =   180
         TabIndex        =   80
         Top             =   720
         Width           =   1350
      End
      Begin VB.CheckBox chk_NoWork_YN 
         BackColor       =   &H00404040&
         Caption         =   "����3"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   2970
         TabIndex        =   79
         Top             =   300
         Width           =   1350
      End
      Begin VB.CheckBox chk_NoWork_YN 
         BackColor       =   &H00404040&
         Caption         =   "����2"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   1560
         TabIndex        =   78
         Top             =   300
         Width           =   1350
      End
      Begin VB.CheckBox chk_NoWork_YN 
         BackColor       =   &H00404040&
         Caption         =   "����1"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   180
         TabIndex        =   77
         Top             =   300
         Width           =   1350
      End
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H00C0C0C0&
      Caption         =   "15. �湮���� �湮�� ���� ���� "
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   15495
      TabIndex        =   69
      Top             =   1170
      Visible         =   0   'False
      Width           =   4470
      Begin VB.CheckBox chk_Guest_YN 
         BackColor       =   &H00C0C0C0&
         Caption         =   "����1"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   75
         Top             =   390
         Width           =   1350
      End
      Begin VB.CheckBox chk_Guest_YN 
         BackColor       =   &H00C0C0C0&
         Caption         =   "����2"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   1590
         TabIndex        =   74
         Top             =   390
         Width           =   1350
      End
      Begin VB.CheckBox chk_Guest_YN 
         BackColor       =   &H00C0C0C0&
         Caption         =   "����3"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   2940
         TabIndex        =   73
         Top             =   390
         Width           =   1350
      End
      Begin VB.CheckBox chk_Guest_YN 
         BackColor       =   &H00C0C0C0&
         Caption         =   "����4"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   240
         TabIndex        =   72
         Top             =   750
         Width           =   1350
      End
      Begin VB.CheckBox chk_Guest_YN 
         BackColor       =   &H00C0C0C0&
         Caption         =   "����5"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   1590
         TabIndex        =   71
         Top             =   750
         Width           =   1350
      End
      Begin VB.CheckBox chk_Guest_YN 
         BackColor       =   &H00C0C0C0&
         Caption         =   "����6"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   2940
         TabIndex        =   70
         Top             =   750
         Width           =   1350
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00404040&
      Caption         =   "��������� ����(�����ͺ��̽�)"
      ForeColor       =   &H00FFFFFF&
      Height          =   1065
      Left            =   180
      TabIndex        =   65
      Top             =   9735
      Visible         =   0   'False
      Width           =   4470
      Begin Threed.SSCommand cmd_Del_Button 
         Height          =   540
         Left            =   2895
         TabIndex        =   66
         Top             =   345
         Width           =   1170
         _Version        =   65536
         _ExtentX        =   2064
         _ExtentY        =   952
         _StockProps     =   78
         Caption         =   "�ϰ� ����"
         ForeColor       =   255
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
         Picture         =   "FrmExtend.frx":0000
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   345
         Left            =   300
         TabIndex        =   67
         Top             =   420
         Width           =   2040
         _ExtentX        =   3598
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
         Format          =   139067392
         CurrentDate     =   36927
      End
      Begin VB.Label lbl_option 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "����, ���� �������ڷ�"
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
         Height          =   285
         Index           =   2
         Left            =   2400
         TabIndex        =   68
         Top             =   450
         Width           =   2220
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00404040&
      Caption         =   " 4. ���ܱ� �ڵ����� [��������] "
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1125
      Left            =   180
      TabIndex        =   58
      Top             =   4470
      Width           =   4470
      Begin VB.CheckBox chk_Taxi_YN 
         BackColor       =   &H00404040&
         Caption         =   "����1"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   64
         Top             =   330
         Width           =   1350
      End
      Begin VB.CheckBox chk_Taxi_YN 
         BackColor       =   &H00404040&
         Caption         =   "����2"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         Left            =   1620
         TabIndex        =   63
         Top             =   330
         Width           =   1350
      End
      Begin VB.CheckBox chk_Taxi_YN 
         BackColor       =   &H00404040&
         Caption         =   "����3"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   2
         Left            =   3000
         TabIndex        =   62
         Top             =   330
         Width           =   1350
      End
      Begin VB.CheckBox chk_Taxi_YN 
         BackColor       =   &H00404040&
         Caption         =   "����4"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   3
         Left            =   240
         TabIndex        =   61
         Top             =   660
         Width           =   1350
      End
      Begin VB.CheckBox chk_Taxi_YN 
         BackColor       =   &H00404040&
         Caption         =   "����5"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   4
         Left            =   1620
         TabIndex        =   60
         Top             =   660
         Width           =   1350
      End
      Begin VB.CheckBox chk_Taxi_YN 
         BackColor       =   &H00404040&
         Caption         =   "����6"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   5
         Left            =   3000
         TabIndex        =   59
         Top             =   660
         Width           =   1350
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00404040&
      Caption         =   " 9. ���� ���� ���� "
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   810
      Left            =   4815
      TabIndex        =   52
      Top             =   1740
      Width           =   4470
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
         ItemData        =   "FrmExtend.frx":0351
         Left            =   240
         List            =   "FrmExtend.frx":0353
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   53
         Top             =   315
         Width           =   2145
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00404040&
      Caption         =   " 10. ���� ���Ͽ��� ���� "
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   810
      Left            =   4815
      TabIndex        =   50
      Top             =   2715
      Width           =   4470
      Begin VB.CheckBox chk_Week_YN 
         BackColor       =   &H00404040&
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   240
         TabIndex        =   51
         Top             =   300
         Width           =   1185
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00404040&
      Caption         =   " 11. ������ ��� �����Ⱓ "
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1065
      Left            =   4815
      TabIndex        =   42
      Top             =   3720
      Width           =   4470
      Begin VB.TextBox txt_using_date 
         Height          =   315
         Left            =   210
         TabIndex        =   43
         Text            =   "99"
         Top             =   570
         Width           =   1000
      End
      Begin VB.Label Label11 
         BackColor       =   &H00404040&
         Caption         =   "1= 1����, 2= 2����, 99= 9999��12��31��"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   180
         TabIndex        =   44
         Top             =   300
         Width           =   4125
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00404040&
      Caption         =   " 12. ����� ������ "
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1065
      Left            =   4815
      TabIndex        =   29
      Top             =   4920
      Width           =   4470
      Begin VB.TextBox Text_EndDate 
         Height          =   315
         Left            =   210
         TabIndex        =   30
         Text            =   "99"
         Top             =   570
         Width           =   1000
      End
      Begin VB.Label Label10 
         BackColor       =   &H00404040&
         Caption         =   "1= 1����, 2= 2����, 99= 9999��12��31��"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   180
         TabIndex        =   31
         Top             =   330
         Width           =   4125
      End
   End
   Begin VB.Frame Ÿ�� 
      BackColor       =   &H00404040&
      Caption         =   " 1. �̿��� Ÿ�� "
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   885
      Left            =   180
      TabIndex        =   28
      Top             =   870
      Width           =   4470
      Begin VB.ComboBox cmb_UserType 
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
         ItemData        =   "FrmExtend.frx":0355
         Left            =   210
         List            =   "FrmExtend.frx":0357
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   54
         Top             =   390
         Width           =   2475
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "2. �����ͺ��̽� ����"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1065
      Left            =   180
      TabIndex        =   16
      Top             =   1905
      Width           =   4470
      Begin VB.TextBox Text_DB_IP 
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   240
         TabIndex        =   18
         Text            =   "192.168.100.200"
         Top             =   600
         Width           =   1725
      End
      Begin VB.TextBox Text_DB_Name 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2580
         TabIndex        =   17
         Text            =   "jwt_anps"
         Top             =   600
         Width           =   1200
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
         Caption         =   "�����ͺ��̽� IP �ּ�"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   330
         Width           =   2145
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404040&
         Caption         =   "�����ͺ��̽� ����"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   2580
         TabIndex        =   19
         Top             =   330
         Width           =   1605
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      Caption         =   "7. ����� �������� ��� Ȱ��ȭ(������Ʈ) "
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1125
      Left            =   180
      TabIndex        =   14
      Top             =   8385
      Width           =   4470
      Begin VB.CheckBox chk_BlackList_YN 
         BackColor       =   &H00404040&
         Caption         =   "����6"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   5
         Left            =   2940
         TabIndex        =   49
         Top             =   750
         Width           =   1350
      End
      Begin VB.CheckBox chk_BlackList_YN 
         BackColor       =   &H00404040&
         Caption         =   "����5"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   4
         Left            =   1560
         TabIndex        =   48
         Top             =   750
         Width           =   1350
      End
      Begin VB.CheckBox chk_BlackList_YN 
         BackColor       =   &H00404040&
         Caption         =   "����4"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   3
         Left            =   240
         TabIndex        =   47
         Top             =   750
         Width           =   1350
      End
      Begin VB.CheckBox chk_BlackList_YN 
         BackColor       =   &H00404040&
         Caption         =   "����3"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   2
         Left            =   2940
         TabIndex        =   46
         Top             =   390
         Width           =   1350
      End
      Begin VB.CheckBox chk_BlackList_YN 
         BackColor       =   &H00404040&
         Caption         =   "����2"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         Left            =   1590
         TabIndex        =   45
         Top             =   390
         Width           =   1350
      End
      Begin VB.CheckBox chk_BlackList_YN 
         BackColor       =   &H00404040&
         Caption         =   "����1"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Top             =   390
         Width           =   1350
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00404040&
      Caption         =   " 8. �ѱ� ���ν� ���͸� ���� "
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   4845
      TabIndex        =   12
      Top             =   870
      Width           =   4470
      Begin VB.CheckBox chk_MissMatch_YN 
         BackColor       =   &H00404040&
         Caption         =   "���"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   240
         TabIndex        =   13
         Top             =   300
         Width           =   825
      End
   End
   Begin VB.CheckBox chk_MVR_YN 
      BackColor       =   &H00404040&
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   8520
      TabIndex        =   7
      Top             =   8310
      Width           =   735
   End
   Begin VB.Frame frmMVR 
      BackColor       =   &H00404040&
      Caption         =   " 14. MVR ���� "
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1065
      Left            =   4785
      TabIndex        =   6
      Top             =   8430
      Width           =   4470
      Begin VB.TextBox Text_MVR_Port 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1860
         TabIndex        =   9
         Text            =   "18498"
         Top             =   540
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox Text_MVR_IP 
         Height          =   315
         Left            =   180
         TabIndex        =   8
         Text            =   "192.168.100.200"
         Top             =   540
         Width           =   1400
      End
      Begin VB.Label Label7 
         BackColor       =   &H00404040&
         Caption         =   "��Ʈ"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1890
         TabIndex        =   11
         Top             =   270
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Label6 
         BackColor       =   &H00404040&
         Caption         =   "������"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   210
         TabIndex        =   10
         Top             =   270
         Width           =   765
      End
   End
   Begin VB.CheckBox chk_HomeNet_YN 
      BackColor       =   &H00404040&
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   8580
      TabIndex        =   5
      Top             =   6090
      Width           =   675
   End
   Begin VB.Frame frmHomeNet 
      BackColor       =   &H00404040&
      Caption         =   " 13. �����뺸 ���� "
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1965
      Left            =   4815
      TabIndex        =   4
      Top             =   6210
      Width           =   4470
      Begin VB.CheckBox chk_MissMatch_HomeNet_YN 
         BackColor       =   &H00404040&
         Caption         =   "�ѱ����͸� �����뺸"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2400
         TabIndex        =   87
         Top             =   1530
         Width           =   1980
      End
      Begin VB.ComboBox cmb_HomeNet 
         Height          =   300
         ItemData        =   "FrmExtend.frx":0359
         Left            =   2400
         List            =   "FrmExtend.frx":035B
         Style           =   2  '��Ӵٿ� ���
         TabIndex        =   55
         Top             =   615
         Width           =   1395
      End
      Begin VB.TextBox Text_HomeNet_Port 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1650
         TabIndex        =   25
         Text            =   "18497"
         Top             =   615
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.CommandButton cmd_HomeTest 
         Caption         =   "�����뺸 �׽�Ʈ"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   24
         Top             =   1005
         Width           =   1410
      End
      Begin VB.TextBox txt_Dong 
         Height          =   315
         Left            =   225
         TabIndex        =   23
         Text            =   "102"
         Top             =   1020
         Width           =   630
      End
      Begin VB.TextBox txt_Ho 
         Height          =   315
         Left            =   1215
         TabIndex        =   22
         Text            =   "101"
         Top             =   1020
         Width           =   630
      End
      Begin VB.TextBox Text_HomeNet_IP 
         Height          =   315
         Left            =   210
         TabIndex        =   21
         Text            =   "192.168.100.200"
         Top             =   615
         Width           =   1400
      End
      Begin VB.Label Label8 
         BackColor       =   &H00404040&
         Caption         =   "ȣ"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1890
         TabIndex        =   57
         Top             =   1050
         Width           =   255
      End
      Begin VB.Label Label4 
         BackColor       =   &H00404040&
         Caption         =   "��"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   900
         TabIndex        =   56
         Top             =   1050
         Width           =   255
      End
      Begin VB.Label Label3 
         BackColor       =   &H00404040&
         Caption         =   "������"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   345
         Width           =   885
      End
      Begin VB.Label Label5 
         BackColor       =   &H00404040&
         Caption         =   "��Ʈ"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1650
         TabIndex        =   26
         Top             =   345
         Visible         =   0   'False
         Width           =   825
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00404040&
      Caption         =   " 6. ���ܱ� �ڵ����� [���ν�����] "
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1125
      Left            =   180
      TabIndex        =   2
      Top             =   7080
      Width           =   4470
      Begin VB.CheckBox chk_NoRecOpen_YN 
         BackColor       =   &H00404040&
         Caption         =   "����6"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Index           =   5
         Left            =   2910
         TabIndex        =   36
         Top             =   750
         Width           =   1350
      End
      Begin VB.CheckBox chk_NoRecOpen_YN 
         BackColor       =   &H00404040&
         Caption         =   "����5"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Index           =   4
         Left            =   1560
         TabIndex        =   35
         Top             =   750
         Width           =   1350
      End
      Begin VB.CheckBox chk_NoRecOpen_YN 
         BackColor       =   &H00404040&
         Caption         =   "����4"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Index           =   3
         Left            =   180
         TabIndex        =   34
         Top             =   750
         Width           =   1350
      End
      Begin VB.CheckBox chk_NoRecOpen_YN 
         BackColor       =   &H00404040&
         Caption         =   "����3"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Index           =   2
         Left            =   2910
         TabIndex        =   33
         Top             =   330
         Width           =   1350
      End
      Begin VB.CheckBox chk_NoRecOpen_YN 
         BackColor       =   &H00404040&
         Caption         =   "����2"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Index           =   1
         Left            =   1560
         TabIndex        =   32
         Top             =   330
         Width           =   1350
      End
      Begin VB.CheckBox chk_NoRecOpen_YN 
         BackColor       =   &H00404040&
         Caption         =   "����1"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Index           =   0
         Left            =   180
         TabIndex        =   3
         Top             =   330
         Width           =   1350
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00404040&
      Caption         =   " 5. ���ܱ� �ڵ����� [�湮����] "
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1125
      Left            =   180
      TabIndex        =   0
      Top             =   5775
      Width           =   4470
      Begin VB.CheckBox chk_FreePassLane_YN 
         BackColor       =   &H00404040&
         Caption         =   "����6"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   5
         Left            =   2970
         TabIndex        =   41
         Top             =   720
         Width           =   1350
      End
      Begin VB.CheckBox chk_FreePassLane_YN 
         BackColor       =   &H00404040&
         Caption         =   "����5"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   4
         Left            =   1560
         TabIndex        =   40
         Top             =   720
         Width           =   1350
      End
      Begin VB.CheckBox chk_FreePassLane_YN 
         BackColor       =   &H00404040&
         Caption         =   "����4"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   3
         Left            =   180
         TabIndex        =   39
         Top             =   720
         Width           =   1350
      End
      Begin VB.CheckBox chk_FreePassLane_YN 
         BackColor       =   &H00404040&
         Caption         =   "����3"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   2970
         TabIndex        =   38
         Top             =   300
         Width           =   1350
      End
      Begin VB.CheckBox chk_FreePassLane_YN 
         BackColor       =   &H00404040&
         Caption         =   "����2"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   1
         Left            =   1560
         TabIndex        =   37
         Top             =   300
         Width           =   1350
      End
      Begin VB.CheckBox chk_FreePassLane_YN 
         BackColor       =   &H00404040&
         Caption         =   "����1"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Top             =   300
         Width           =   1350
      End
   End
   Begin Threed.SSCommand cmdCancel 
      Height          =   585
      Left            =   12660
      TabIndex        =   88
      Top             =   195
      Width           =   1320
      _Version        =   65536
      _ExtentX        =   2328
      _ExtentY        =   1032
      _StockProps     =   78
      Caption         =   "�� ��"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�������"
         Size            =   12.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "FrmExtend.frx":035D
   End
   Begin Threed.SSCommand cmdOK 
      Height          =   585
      Left            =   11205
      TabIndex        =   89
      Top             =   195
      Width           =   1320
      _Version        =   65536
      _ExtentX        =   2328
      _ExtentY        =   1032
      _StockProps     =   78
      Caption         =   "�� ��"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�������"
         Size            =   12.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "FrmExtend.frx":06AE
   End
   Begin VB.Label Label9 
      BackColor       =   &H00404040&
      Caption         =   "���γ��� ȯ�漳��"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   14.25
         Charset         =   129
         Weight          =   600
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   255
      TabIndex        =   90
      Top             =   270
      Width           =   4155
   End
End
Attribute VB_Name = "FrmExtend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



'Const NAME_COLUMN = 0
'Const TYPE_COLUMN = 1
'Const SIZE_COLUMN = 2
'Const DATE_COLUMN = 3
'
'Private Sub mnuFileClose_Click()
'  MsgBox "�ݱ� �ڵ带 �ۼ��Ͻʽÿ�!"
'End Sub
'
'Private Sub mnuFileExit_Click()
'  '���� ��ε��մϴ�.
'  Unload Me
'End Sub
'
'Private Sub mnuFileNew_Click()
'  MsgBox "�� ���� �ڵ带 �ۼ��Ͻʽÿ�!"
'End Sub
'
'Private Sub mnuFileOpen_Click()
'  MsgBox "���� �ڵ带 �ۼ��Ͻʽÿ�!"
'End Sub
'
'Private Sub mnuFilePrint_Click()
'  MsgBox "�μ� �ڵ带 �ۼ��Ͻʽÿ�!"
'End Sub
'
'Private Sub mnuFilePrintPreview_Click()
'  MsgBox "�μ� �̸����� �ڵ带 �ۼ��Ͻʽÿ�!"
'End Sub
'
'Private Sub mnuFilePrintSetup_Click()
'  MsgBox "������ ���� �ڵ带 �ۼ��Ͻʽÿ�!"
'End Sub
'
'Private Sub mnuFileProperties_Click()
'  MsgBox "�Ӽ� �ڵ带 �ۼ��Ͻʽÿ�!"
'End Sub
'
'Private Sub mnuFileSave_Click()
'  MsgBox "���� ���� �ڵ带 �ۼ��Ͻʽÿ�!"
'End Sub
'
'Private Sub mnuFileSaveAll_Click()
'  MsgBox "��� ���� �ڵ带 �ۼ��Ͻʽÿ�!"
'End Sub
'
'Private Sub mnuFileSaveAs_Click()
'  MsgBox "�ٸ� �̸����� ���� �ڵ带 �ۼ��Ͻʽÿ�!"
'End Sub
'
'Private Sub mnuFileSend_Click()
'  MsgBox "������ �ڵ带 �ۼ��Ͻʽÿ�!"
'End Sub
'
'Private Sub mnuVAIByDate_Click()
''  lvListView.SortKey = DATE_COLUMN
'End Sub
'
'Private Sub mnuVAIByName_Click()
''  lvListView.SortKey = NAME_COLUMN
'End Sub
'
'Private Sub mnuVAIBySize_Click()
''  lvListView.SortKey = SIZE_COLUMN
'End Sub
'
'Private Sub mnuVAIByType_Click()
''  lvListView.SortKey = TYPE_COLUMN
'End Sub
'
'Private Sub mnuViewDetails_Click()
''  lvListView.View = lvwReport
'End Sub
'
'Private Sub mnuViewLargeIcons_Click()
''  lvListView.View = lvwIcon
'End Sub
'
'Private Sub mnuViewLineUpIcons_Click()
''  lvListView.Arrange = lvwAutoLeft
'End Sub
'
'Private Sub mnuViewList_Click()
''  lvListView.View = lvwList
'End Sub
'
'Private Sub mnuViewOptions_Click()
''  frmOptions.Show vbModal
'End Sub
'
'Private Sub mnuViewRefresh_Click()
'  MsgBox "���� ��ħ �ڵ带 ���⿡ �νʽÿ�."
'End Sub
'
'Private Sub mnuViewSmallIcons_Click()
''  lvListView.View = lvwSmallIcon
'End Sub
'
'Private Sub mnuViewStatusBar_Click()
'  If mnuViewStatusBar.Checked Then
''    sbStatusBar.Visible = False
'    mnuViewStatusBar.Checked = False
'  Else
''    sbStatusBar.Visible = True
'    mnuViewStatusBar.Checked = True
'  End If
'End Sub
'
'Private Sub mnuViewToolbar_Click()
'  If mnuViewToolbar.Checked Then
''    tbToolBar.Visible = False
'    mnuViewToolbar.Checked = False
'  Else
''    tbToolBar.Visible = True
'    mnuViewToolbar.Checked = True
'  End If
'End Sub
'Private Sub cmdAdd_Click()
'  Dim sTmp As String
'  sTmp = InputBox("�߰��� �� �׸��� �Է��Ͻʽÿ�.")
'  If Len(sTmp) = 0 Then Exit Sub
'  lstItems.AddItem sTmp
'End Sub
'
'Private Sub cmdDelete_Click()
'  If lstItems.ListIndex > -1 Then
'    If MsgBox(lstItems.Text & "(��)�� �����Ͻðڽ��ϱ�?", vbQuestion + vbYesNo) = vbYes Then
'      lstItems.RemoveItem lstItems.ListIndex
'    End If
'  End If
'End Sub
'
'Private Sub cmdUp_Click()
'  On Error Resume Next
'  Dim nItem As Integer
'
'  With lstItems
'    If .ListIndex < 0 Then Exit Sub
'    nItem = .ListIndex
'    If nItem = 0 Then Exit Sub  'ù° �׸��� ���� �̵��� �� �����ϴ�.
'    '�׸��� ���� �̵��մϴ�.
'    .AddItem .Text, nItem - 1
'    '���� �׸��� �����մϴ�.
'    .RemoveItem nItem + 1
'    '��� �̵��� �׸��� �����մϴ�.
'    .Selected(nItem - 1) = True
'  End With
'End Sub
'
'Private Sub cmdDown_Click()
'  On Error Resume Next
'  Dim nItem As Integer
'
'  With lstItems
'    If .ListIndex < 0 Then Exit Sub
'    nItem = .ListIndex
'    If nItem = .ListCount - 1 Then Exit Sub '������ �׸��� �Ʒ��� �̵��� �� �����ϴ�.
'    '�׸��� �Ʒ��� �̵��մϴ�.
'    .AddItem .Text, nItem + 2
'    '���� �׸��� �����մϴ�.
'    .RemoveItem nItem
'    '��� �̵��� �׸��� �����մϴ�.
'    .Selected(nItem + 1) = True
'  End With
'End Sub
'
'Private Sub lstItems_DragDrop(Source As Control, X As Single, Y As Single)
'  Dim i As Integer
'  Dim nID As Integer
'  Dim sTmp As String
'
'  If Source.Name <> "lstItems" Then Exit Sub
'  If lstItems.ListCount = 0 Then Exit Sub
'
'  With lstItems
'    i = (Y \ TextHeight("A")) + .TopIndex
'    If i = .ListIndex Then
'      '�ڽ��� �� �κп� �����ϴ�.
'      Exit Sub
'    End If
'    If i > .ListCount - 1 Then i = .ListCount - 1
'    nID = .ListIndex
'    sTmp = .Text
'    If (nID > -1) Then
'      sTmp = .Text
'      .RemoveItem nID
'      .AddItem sTmp, i
'      .ListIndex = .NewIndex
'    End If
'  End With
'  SetListButtons
'End Sub
'
'Sub lstItems_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'  If Button = vbLeftButton Then lstItems.Drag
'End Sub
'
'Private Sub lstItems_Click()
'  SetListButtons
'End Sub
'
'Sub SetListButtons()
'  Dim i As Integer
'  i = lstItems.ListIndex
'  '�̵� ������ ���¸� �����մϴ�.
'  cmdUp.Enabled = (i > 0)
'  cmdDown.Enabled = ((i > -1) And (i < (lstItems.ListCount - 1)))
'  cmdDelete.Enabled = (i > -1)
'End Sub




Private Sub chk_HomeNet_YN_Click()
    If chk_HomeNet_YN.value = 1 Then
        frmHomeNet.Enabled = True
    Else
        frmHomeNet.Enabled = False
    End If

End Sub

Private Sub chk_MVR_YN_Click()
    If chk_MVR_YN.value = 1 Then
        frmMVR.Enabled = True
    Else
        frmMVR.Enabled = False
    End If
    
End Sub



Private Sub chk_Week_YN_Click()
    If chk_Week_YN.value = 1 Then
        If (cmb_Rotation.ListCount > 0) Then
            cmb_Rotation.ListIndex = 0
        End If
        cmb_Rotation.Enabled = False
    Else
        cmb_Rotation.Enabled = True
    End If
End Sub



Private Sub cmb_Rotation_Click()
    If (cmb_Rotation.ListIndex <> 0) Then
        chk_Week_YN.value = 0
        chk_Week_YN.Enabled = False
    Else
        chk_Week_YN.Enabled = True
    End If
End Sub



Private Sub cmd_Del_Button_Click()
    Dim sQry As String
    Dim sLog As String

    MBox.Label3.Caption = Format(DTPicker1, "yyyy-mm-dd") & " ����"
    MBox.Label1.Caption = "���� ������������ �����Ͻðڽ��ϱ�?"
    MBox.Label2.Caption = "���������� ���� ����"
    MBox.Show 1
    If (Glo_MsgRet = True) Then

        sQry = "Delete From tb_inout Where PASS_DATE <= '" & Format(DTPicker1, "yyyy-mm-dd") & " 23:59:59 999" & "' "
        adoConn.Execute sQry
        
        sLog = Format(DTPicker1, "yyyy-mm-dd") & " �����Ͽ� ���� �������ڷ� ����"
        adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('����������', 'HOST','" & sLog & "','�ڷ����'," & 0 & "," '" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"

        Call DataLogger(Format(DTPicker1, "yyyy-mm-dd") & " �����Ͽ� ���� �������ڷ� ����")
    Else
    End If
    
End Sub


Private Sub cmd_HomeTest_Click()
    HomeNet_Dong = txt_Dong.text
    HomeNet_Ho = txt_Ho.text
    HomeNet_CarNo = "����01��1234"
    
    HomeNet_Str = HomeNet_Dong & HomeNet_Ho & HomeNet_CarNo
    

    If (FrmTcpServer.HomeSock.State = sckClosed) Then

        If (HomeNet_IP <> "" And HomeNet_Port > 0) Then
        
            FrmTcpServer.HomeSock.Protocol = sckUDPProtocol
            FrmTcpServer.HomeSock.RemoteHost = HomeNet_IP
            FrmTcpServer.HomeSock.RemotePort = HomeNet_Port
    
            FrmTcpServer.HomeSock.SendData (HomeNet_Str)
            Call DataLogger("[HomeNet UDP ����]  IP = " & HomeNet_IP & "    PORT = " & HomeNet_Port & "      DATA = " & HomeNet_Str)
            
        Else
            Call DataLogger("[HomeNet UDP ����]  HomeNet IP �� HomeNet Port Ȯ�� �� �������ּ���")
        End If
    Else
        FrmTcpServer.HomeSock.SendData (HomeNet_Str)
        Call DataLogger("[HomeNet UDP ����]  IP = " & HomeNet_IP & "    PORT = " & HomeNet_Port & "      DATA = " & HomeNet_Str)
    End If
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim ECHO As ICMP_ECHO_REPLY
    Dim sip_pos As Integer
    Dim eip_pos As Integer
    Dim old_dbip As String
    
    Dim sname_pos As Integer
    Dim ename_pos As Integer
    Dim old_dbname As String
    
    Dim TempAdoConn As String
    
    Dim i As Integer

On Error GoTo Err_p

    cmdOK.Enabled = False
    cmdCancel.Enabled = False
    
    
    '�����ͺ��̽� ����
    sip_pos = InStr(UCase(AdoConn_Str), "SERVER=") + Len("SERVER=")
    eip_pos = InStr(UCase(AdoConn_Str), "DATABASE=")
    old_dbip = Mid(AdoConn_Str, sip_pos, eip_pos - sip_pos - 1)
    'AdoConn_Str = Replace(AdoConn_Str, old_dbip, Text_DB_IP)
    

    sname_pos = InStr(UCase(AdoConn_Str), "DATABASE=") + Len("DATABASE=")
    ename_pos = InStr(UCase(AdoConn_Str), "UID=")
    old_dbname = Mid(AdoConn_Str, sname_pos, ename_pos - sname_pos - 1)
    'AdoConn_Str = Replace(AdoConn_Str, old_dbname, Text_DB_Name)


    If (old_dbip <> Text_DB_IP Or old_dbname <> Text_DB_Name) Then
        '1. Ping���� Ȯ��
        Call Ping(Text_DB_IP, ECHO)
        If Left$(ECHO.Data, 1) <> Chr$(0) Then
            
            '2. ������ DB IP�� Name�� �̿��� ���� �׽�Ʈ
            TempAdoConn = AdoConn_Str
            TempAdoConn = Replace(TempAdoConn, old_dbip, Text_DB_IP)
            TempAdoConn = Replace(TempAdoConn, old_dbname, Text_DB_Name)
    
            If (adoTemp.State = adStateOpen) Then
                Call DataBaseClose(adoTemp)
            End If
            
            i = 0
            Do While DataBaseOpenTemp(adoTemp, TempAdoConn) = False
                Call DataLogger("[FrmExtend] DB Temp Connection Failure..!!")
                Call Delay_Time(1)
                i = i + 1
                If i > 3 Then
                    Call MsgBox("DataBase Name Ȯ���� �ٽ� �������ּ���", vbInformation Or vbMsgBoxSetForeground, "DataBase Name ����")
                    Exit Do
                End If
            Loop
            
            If (adoTemp.State = adStateOpen) Then
            
                '3. ���ο� DB IP�� Name���� ��ü
                Call DataBaseClose(adoTemp)
                
                AdoConn_Str = TempAdoConn
                If (adoConn.State = adStateOpen) Then
                    Call DataBaseClose(adoConn)
                End If
                
                Do While DataBaseOpen(adoConn) = False
                    Call DataLogger("[FrmExtend] DB Connection Failure..!!")
                    Call Delay_Time(1)
                    i = i + 1
                    If i > 3 Then
                        Call MsgBox("DataBase IP �ּ�, Name Ȯ���� �ٽ� �������ּ���", vbInformation Or vbMsgBoxSetForeground, "DataBase ���� ����")
                        'End
                        'Return
                        Exit Do
                    End If
                Loop
            End If
        Else
            Call MsgBox("DataBase IP�ּ� Ȯ���� �ٽ� �������ּ���", vbInformation Or vbMsgBoxSetForeground, "DataBase IP�ּ� ����")
            Call DataLogger("[FrmExtend Save]    Ping Test Failure...!!")
            Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & "[FrmExtend]" & "Ping Test Failure...!!")
        End If
    End If

    DB_Server_IP = Text_DB_IP
    DB_Server_Port = 3306
    
    
    
    
    
    
    
    
    Glo_EndDate = Val(Text_EndDate)

    Glo_User_Type = cmb_UserType.text
    

    If (chk_BlackList_YN(0).value = 1) Then
        Glo_BlackList1_YN = "Y"
    Else
        Glo_BlackList1_YN = "N"
    End If
    If (chk_BlackList_YN(1).value = 1) Then
        Glo_BlackList2_YN = "Y"
    Else
        Glo_BlackList2_YN = "N"
    End If
    If (chk_BlackList_YN(2).value = 1) Then
        Glo_BlackList3_YN = "Y"
    Else
        Glo_BlackList3_YN = "N"
    End If
    If (chk_BlackList_YN(3).value = 1) Then
        Glo_BlackList4_YN = "Y"
    Else
        Glo_BlackList4_YN = "N"
    End If
    If (chk_BlackList_YN(4).value = 1) Then
        Glo_BlackList5_YN = "Y"
    Else
        Glo_BlackList5_YN = "N"
    End If
    If (chk_BlackList_YN(5).value = 1) Then
        Glo_BlackList6_YN = "Y"
    Else
        Glo_BlackList6_YN = "N"
    End If
    

    
    If (chk_FreePassLane_YN(0).value = 1) Then
        Glo_FreePassLane1_YN = "Y"
    Else
        Glo_FreePassLane1_YN = "N"
    End If
    If (chk_FreePassLane_YN(1).value = 1) Then
        Glo_FreePassLane2_YN = "Y"
    Else
        Glo_FreePassLane2_YN = "N"
    End If
    If (chk_FreePassLane_YN(2).value = 1) Then
        Glo_FreePassLane3_YN = "Y"
    Else
        Glo_FreePassLane3_YN = "N"
    End If
    If (chk_FreePassLane_YN(3).value = 1) Then
        Glo_FreePassLane4_YN = "Y"
    Else
        Glo_FreePassLane4_YN = "N"
    End If
    If (chk_FreePassLane_YN(4).value = 1) Then
        Glo_FreePassLane5_YN = "Y"
    Else
        Glo_FreePassLane5_YN = "N"
    End If
    If (chk_FreePassLane_YN(5).value = 1) Then
        Glo_FreePassLane6_YN = "Y"
    Else
        Glo_FreePassLane6_YN = "N"
    End If
    
    Call FrmG1.ReDraw("FreePass", 0, chk_FreePassLane_YN(0).value)
    For i = 0 To 1
        Call Jung.ReDraw("FreePass", i, chk_FreePassLane_YN(i).value)
    Next i
    For i = 0 To 4
        Call FrmG4Mini.ReDraw("FreePass", i, chk_FreePassLane_YN(i).value)
    Next i
    
    
    
    If (chk_NoRecOpen_YN(0).value = 1) Then
        Glo_NoRecOpen1_YN = "Y"
    Else
        Glo_NoRecOpen1_YN = "N"
    End If
    If (chk_NoRecOpen_YN(1).value = 1) Then
        Glo_NoRecOpen2_YN = "Y"
    Else
        Glo_NoRecOpen2_YN = "N"
    End If
    If (chk_NoRecOpen_YN(2).value = 1) Then
        Glo_NoRecOpen3_YN = "Y"
    Else
        Glo_NoRecOpen3_YN = "N"
    End If
    If (chk_NoRecOpen_YN(3).value = 1) Then
        Glo_NoRecOpen4_YN = "Y"
    Else
        Glo_NoRecOpen4_YN = "N"
    End If
    If (chk_NoRecOpen_YN(4).value = 1) Then
        Glo_NoRecOpen5_YN = "Y"
    Else
        Glo_NoRecOpen5_YN = "N"
    End If
    If (chk_NoRecOpen_YN(5).value = 1) Then
        Glo_NoRecOpen6_YN = "Y"
    Else
        Glo_NoRecOpen6_YN = "N"
    End If



    '����������
    If (chk_Taxi_YN(0).value = 1) Then
        Glo_TAXI1_YN = "Y"
    Else
        Glo_TAXI1_YN = "N"
    End If
    If (chk_Taxi_YN(1).value = 1) Then
        Glo_TAXI2_YN = "Y"
    Else
        Glo_TAXI2_YN = "N"
    End If
    If (chk_Taxi_YN(2).value = 1) Then
        Glo_TAXI3_YN = "Y"
    Else
        Glo_TAXI3_YN = "N"
    End If
    If (chk_Taxi_YN(3).value = 1) Then
        Glo_TAXI4_YN = "Y"
    Else
        Glo_TAXI4_YN = "N"
    End If
    If (chk_Taxi_YN(4).value = 1) Then
        Glo_TAXI5_YN = "Y"
    Else
        Glo_TAXI5_YN = "N"
    End If
    If (chk_Taxi_YN(5).value = 1) Then
        Glo_TAXI6_YN = "Y"
    Else
        Glo_TAXI6_YN = "N"
    End If
    
    
    Select Case Glo_Screen_No
        Case 1
            Call FrmG1.ReDraw("Taxi", 0, chk_Taxi_YN(0).value)
        Case 2
            Call Jung.ReDraw("Taxi", 0, chk_Taxi_YN(0).value)
            Call Jung.ReDraw("Taxi", 1, chk_Taxi_YN(1).value)
        Case 4
            Call FrmG4Mini.ReDraw("Taxi", 0, chk_Taxi_YN(0).value)
            Call FrmG4Mini.ReDraw("Taxi", 1, chk_Taxi_YN(1).value)
            Call FrmG4Mini.ReDraw("Taxi", 2, chk_Taxi_YN(2).value)
            Call FrmG4Mini.ReDraw("Taxi", 3, chk_Taxi_YN(3).value)
        Case 6
            Call FrmG6_23.ReDraw("Taxi", 0, chk_Taxi_YN(0).value)
            Call FrmG6_23.ReDraw("Taxi", 1, chk_Taxi_YN(1).value)
            Call FrmG6_23.ReDraw("Taxi", 2, chk_Taxi_YN(2).value)
            Call FrmG6_23.ReDraw("Taxi", 3, chk_Taxi_YN(3).value)
            Call FrmG6_23.ReDraw("Taxi", 4, chk_Taxi_YN(4).value)
            Call FrmG6_23.ReDraw("Taxi", 5, chk_Taxi_YN(5).value)
    End Select
    

    '�ڸ���� ���� ����
    If (chk_NoWork_YN(0).value = 1) Then
        Glo_NOWORK1_YN = "Y"
    Else
        Glo_NOWORK1_YN = "N"
    End If
    If (chk_NoWork_YN(1).value = 1) Then
        Glo_NOWORK2_YN = "Y"
    Else
        Glo_NOWORK2_YN = "N"
    End If
    If (chk_NoWork_YN(2).value = 1) Then
        Glo_NOWORK3_YN = "Y"
    Else
        Glo_NOWORK3_YN = "N"
    End If
    If (chk_NoWork_YN(3).value = 1) Then
        Glo_NOWORK4_YN = "Y"
    Else
        Glo_NOWORK4_YN = "N"
    End If
    If (chk_NoWork_YN(4).value = 1) Then
        Glo_NOWORK5_YN = "Y"
    Else
        Glo_NOWORK5_YN = "N"
    End If
    If (chk_NoWork_YN(5).value = 1) Then
        Glo_NOWORK6_YN = "Y"
    Else
        Glo_NOWORK6_YN = "N"
    End If
    
    
    Select Case Glo_Screen_No
        Case 1
            Call FrmG1.ReDraw("NOWORK", 0, chk_NoWork_YN(0).value)
        Case 2
            Call Jung.ReDraw("NOWORK", 0, chk_NoWork_YN(0).value)
            Call Jung.ReDraw("NOWORK", 1, chk_NoWork_YN(1).value)
        Case 4
            Call FrmG4Mini.ReDraw("NOWORK", 0, chk_NoWork_YN(0).value)
            Call FrmG4Mini.ReDraw("NOWORK", 1, chk_NoWork_YN(1).value)
            Call FrmG4Mini.ReDraw("NOWORK", 2, chk_NoWork_YN(2).value)
            Call FrmG4Mini.ReDraw("NOWORK", 3, chk_NoWork_YN(3).value)
        Case 6
            Call FrmG6_23.ReDraw("NOWORK", 0, chk_NoWork_YN(0).value)
            Call FrmG6_23.ReDraw("NOWORK", 1, chk_NoWork_YN(1).value)
            Call FrmG6_23.ReDraw("NOWORK", 2, chk_NoWork_YN(2).value)
            Call FrmG6_23.ReDraw("NOWORK", 3, chk_NoWork_YN(3).value)
            Call FrmG6_23.ReDraw("NOWORK", 4, chk_NoWork_YN(4).value)
            Call FrmG6_23.ReDraw("NOWORK", 5, chk_NoWork_YN(5).value)
    End Select
    
    
    '�ѱ����͸�
    If (chk_MissMatch_YN.value = 1) Then
        MissMatch_YN = "Y"
    Else
        MissMatch_YN = "N"
    End If
    
    
    
    'Ȩ�� ����
    cmd_HomeTest.Enabled = True
    If (chk_HomeNet_YN.value = 1) Then
        HomeNet_YN = "Y"
    Else
        HomeNet_YN = "N"
    End If
    HomeNet_IP = Trim(Text_HomeNet_IP)
    HomeNet_Port = Val(Text_HomeNet_Port)
    Call Put_Ini2("System Config", "HomeNetMode", CStr(cmb_HomeNet.ListIndex + 1), "C:\HomeNet\HomeNet.ini")

    If (HomeNet_YN = "Y") Then
        Call RunHomeNet
    End If
    
    '�ѱ����͸� �����뺸
    If (chk_MissMatch_HomeNet_YN.value = 1) Then
        MissMatch_HomeNet_YN = "Y"
    Else
        MissMatch_HomeNet_YN = "N"
    End If
    Call Put_Ini("System Config", "MissMatch_HomeNet_YN", MissMatch_HomeNet_YN)
    
    
    'MVR ����
    If (chk_MVR_YN.value = 1) Then
        MVR_YN = "Y"
    Else
        MVR_YN = "N"
    End If
    MVR_IP = Trim(Text_MVR_IP)
    MVR_Port = Val(Text_MVR_Port)
    
    Shell ("taskkill /f /im MVR.exe")
    If (MVR_YN = "Y") Then
        If (IsFile("C:\MVR\MVR.exe") = True) Then
            
            Delay_Time (1)
            Shell ("C:\MVR\MVR.exe")
            Delay_Time (2)

            FrmTcpServer.MvrSock.Close
            FrmTcpServer.MvrSock.Protocol = sckUDPProtocol
            FrmTcpServer.MvrSock.RemoteHost = sckUDPProtocol
            FrmTcpServer.MvrSock.RemotePort = MVR_Port
        End If
    End If
    
        
    
    '������
    If chk_Week_YN.value = 1 Then
        Glo_WEEK_YN = "Y"
    Else
        Glo_WEEK_YN = "N"
    End If
    
    
    
    '��������
    Glo_ROTATION = cmb_Rotation.text
    If (Glo_ROTATION = "������") Then
        adoConn.Execute "UPDATE tb_reg SET ROTATION = 'N' "
    Else
        adoConn.Execute "UPDATE tb_reg SET ROTATION = 'Y' "
    End If
    
    '���������� ���� �Ⱓ(db)
    Glo_INOUT_USING_DATE = Val(txt_using_date.text)
    If (Glo_INOUT_USING_DATE < 0 Or Glo_INOUT_USING_DATE > 36) Then
        Glo_INOUT_USING_DATE = 99
        txt_using_date.text = "99"
    End If
    
    '������ ���� ����
    Dim sWebDC As String
    If (chk_WebDC.value = 1) Then
        sWebDC = "Y"
    Else
        sWebDC = "N"
    End If
    Glo_WebDC_YN = sWebDC
    adoConn.Execute "UPDATE tb_config set Content = '" & sWebDC & "' WHERE NAME = 'WebDC'"
    adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('ȯ�漳��', 'HOST','������:" & sWebDC & "','������'," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
    '������ ���� ��
    
    
    '�����湮������� ���� ����
    Dim sGuestCarReg As String
    If (chk_GuestCarReg.value = 1) Then
        sGuestCarReg = "Y"
    Else
        sGuestCarReg = "N"
    End If
    Glo_GuestReg_YN = sGuestCarReg
    adoConn.Execute "UPDATE tb_config set Content = '" & sGuestCarReg & "' WHERE NAME = 'GuestCarReg'"
    adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('ȯ�漳��', 'HOST','�湮�����������:" & sGuestCarReg & "','�������'," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
    '�����湮������� ���� ��
    
    
    
    '����� �˸� ��� ����
    Dim sMobileAlarm As String
    If (chk_MobileAlarm.value = 1) Then
        sMobileAlarm = "Y"
    Else
        sMobileAlarm = "N"
    End If
    Glo_MobileAlarm_YN = sMobileAlarm
    adoConn.Execute "UPDATE tb_config set Content = '" & sMobileAlarm & "' WHERE TITLE = 'MOBILE' AND NAME = 'ALARM'"
    adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('ȯ�漳��', 'HOST','����Ͼ˸����:" & sMobileAlarm & "','�����'," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
    '����� �˸� ��� ��
    
    
    
    Call Put_Ini("System Config", "Conn_Str", AdoConn_Str)
    Call Put_Ini("System Config", "END_Date", CStr(Glo_EndDate))
    Call Put_Ini("System Config", "User_Type", Glo_User_Type)
    
    
    Call Put_Ini("System Config", "NOWORK1_YN", Glo_NOWORK1_YN)
    Call Put_Ini("System Config", "NOWORK2_YN", Glo_NOWORK2_YN)
    Call Put_Ini("System Config", "NOWORK3_YN", Glo_NOWORK3_YN)
    Call Put_Ini("System Config", "NOWORK4_YN", Glo_NOWORK4_YN)
    Call Put_Ini("System Config", "NOWORK5_YN", Glo_NOWORK5_YN)
    Call Put_Ini("System Config", "NOWORK6_YN", Glo_NOWORK6_YN)
    
    Call Put_Ini("System Config", "FreePassLane1_YN", Glo_FreePassLane1_YN)
    Call Put_Ini("System Config", "FreePassLane2_YN", Glo_FreePassLane2_YN)
    Call Put_Ini("System Config", "FreePassLane3_YN", Glo_FreePassLane3_YN)
    Call Put_Ini("System Config", "FreePassLane4_YN", Glo_FreePassLane4_YN)
    Call Put_Ini("System Config", "FreePassLane5_YN", Glo_FreePassLane5_YN)
    Call Put_Ini("System Config", "FreePassLane6_YN", Glo_FreePassLane6_YN)
    
    Call Put_Ini("System Config", "TAXI1_YN", Glo_TAXI1_YN)
    Call Put_Ini("System Config", "TAXI2_YN", Glo_TAXI2_YN)
    Call Put_Ini("System Config", "TAXI3_YN", Glo_TAXI3_YN)
    Call Put_Ini("System Config", "TAXI4_YN", Glo_TAXI4_YN)
    Call Put_Ini("System Config", "TAXI5_YN", Glo_TAXI5_YN)
    Call Put_Ini("System Config", "TAXI6_YN", Glo_TAXI6_YN)
    
    Call Put_Ini("System Config", "NoRecOpen1_YN", Glo_NoRecOpen1_YN)
    Call Put_Ini("System Config", "NoRecOpen2_YN", Glo_NoRecOpen2_YN)
    Call Put_Ini("System Config", "NoRecOpen3_YN", Glo_NoRecOpen3_YN)
    Call Put_Ini("System Config", "NoRecOpen4_YN", Glo_NoRecOpen4_YN)
    Call Put_Ini("System Config", "NoRecOpen5_YN", Glo_NoRecOpen5_YN)
    Call Put_Ini("System Config", "NoRecOpen6_YN", Glo_NoRecOpen6_YN)
    
    Call Put_Ini("System Config", "BlackList1_YN", Glo_BlackList1_YN)
    Call Put_Ini("System Config", "BlackList2_YN", Glo_BlackList2_YN)
    Call Put_Ini("System Config", "BlackList3_YN", Glo_BlackList3_YN)
    Call Put_Ini("System Config", "BlackList4_YN", Glo_BlackList4_YN)
    Call Put_Ini("System Config", "BlackList5_YN", Glo_BlackList5_YN)
    Call Put_Ini("System Config", "BlackList6_YN", Glo_BlackList6_YN)
    
    Call Put_Ini("System Config", "GUEST1_YN", Glo_GUEST_LANE1_YN)
    Call Put_Ini("System Config", "GUEST2_YN", Glo_GUEST_LANE2_YN)
    Call Put_Ini("System Config", "GUEST3_YN", Glo_GUEST_LANE3_YN)
    Call Put_Ini("System Config", "GUEST4_YN", Glo_GUEST_LANE4_YN)
    Call Put_Ini("System Config", "GUEST5_YN", Glo_GUEST_LANE5_YN)
    Call Put_Ini("System Config", "GUEST6_YN", Glo_GUEST_LANE6_YN)
    
    
'    Call Put_Ini("System Config", "TAXI_YN", TAXI_YN)
    'Call Put_Ini("System Config", "TAXI_IN_YN", Glo_TAXI_IN_YN)
    'Call Put_Ini("System Config", "TAXI_OUT_YN", Glo_TAXI_OUT_YN)
    
    Call Put_Ini("System Config", "MissMatch_YN", MissMatch_YN)
    Call Put_Ini("System Config", "HomeNet_YN", HomeNet_YN)
    Call Put_Ini("System Config", "HomeNet_IP", HomeNet_IP)
    Call Put_Ini("System Config", "HomeNet_Port", CStr(HomeNet_Port))
    Call Put_Ini("System Config", "MVR_YN", MVR_YN)
    Call Put_Ini("System Config", "MVR_IP", MVR_IP)
    Call Put_Ini("System Config", "MVR_Port", CStr(MVR_Port))
    
    Call Put_Ini("System Config", "INOUT_USING_DATE", Val(txt_using_date.text))
    
    Call Put_Ini("System Config", "WEEK_YN", Glo_WEEK_YN)
    Call Put_Ini("System Config", "ROTATION", Glo_ROTATION)
    
    
    Call DataLogger("[Extend Config] Ȯ�� ȯ�漳�� ����")
'    Me.Hide


    cmdOK.Enabled = True
    cmdCancel.Enabled = True
    
    Unload Me
    Exit Sub
    
Err_p:
    Call DataLogger(" [Extend Config Save Error " & Err.Description)
End Sub


Private Sub Form_Load()

    Dim sip_pos As Integer
    Dim eip_pos As Integer
    Dim old_dbip As String
    
    Dim sname_pos As Integer
    Dim ename_pos As Integer
    Dim old_dbname As String
    
    Dim iHomeNetNo As Integer
    
    Dim i As Integer
    
On Error GoTo Err_p
    
    Left = (Screen.width - width) / 2   ' ���� ���η� �߾ӿ� �����ϴ�.
    Top = (Screen.height - height) / 2   ' ���� ���η� �߾ӿ� �����ϴ�.
    



    sip_pos = InStr(UCase(AdoConn_Str), "SERVER=") + Len("SERVER=")
    eip_pos = InStr(UCase(AdoConn_Str), "DATABASE=")
    old_dbip = Mid(AdoConn_Str, sip_pos, eip_pos - sip_pos - 1)
    
    sname_pos = InStr(UCase(AdoConn_Str), "DATABASE=") + Len("DATABASE=")
    ename_pos = InStr(UCase(AdoConn_Str), "UID=")
    old_dbname = Mid(AdoConn_Str, sname_pos, ename_pos - sname_pos - 1)

    Text_DB_IP.text = old_dbip
    Text_DB_Name.text = old_dbname
    
    
    
    '����Ÿ��
    With cmb_UserType
        .AddItem "����1/����2"
        .AddItem "��/ȣ��"
    End With
    cmb_UserType.text = Glo_User_Type
    

    
    Text_EndDate = Glo_EndDate
    
   

    '���νı��
    For i = 0 To MAX_LANE_COUNT - 1
        chk_NoRecOpen_YN(i).Caption = "�̻��": chk_NoRecOpen_YN(i).Enabled = False
    Next i
    If (LANE1_YN = "Y") Then chk_NoRecOpen_YN(0).Caption = LANE1_Name: chk_NoRecOpen_YN(0).Enabled = True
    If (LANE2_YN = "Y") Then chk_NoRecOpen_YN(1).Caption = LANE2_Name: chk_NoRecOpen_YN(1).Enabled = True
    If (LANE3_YN = "Y") Then chk_NoRecOpen_YN(2).Caption = LANE3_Name: chk_NoRecOpen_YN(2).Enabled = True
    If (LANE4_YN = "Y") Then chk_NoRecOpen_YN(3).Caption = LANE4_Name: chk_NoRecOpen_YN(3).Enabled = True
    If (LANE5_YN = "Y") Then chk_NoRecOpen_YN(4).Caption = LANE5_Name: chk_NoRecOpen_YN(4).Enabled = True
    If (LANE6_YN = "Y") Then chk_NoRecOpen_YN(5).Caption = LANE6_Name: chk_NoRecOpen_YN(5).Enabled = True

    If (Glo_NoRecOpen1_YN = "Y") Then
        chk_NoRecOpen_YN(0).value = 1
    Else
        chk_NoRecOpen_YN(0).value = 0
    End If
    If (Glo_NoRecOpen2_YN = "Y") Then
        chk_NoRecOpen_YN(1).value = 1
    Else
        chk_NoRecOpen_YN(1).value = 0
    End If
    If (Glo_NoRecOpen3_YN = "Y") Then
        chk_NoRecOpen_YN(2).value = 1
    Else
        chk_NoRecOpen_YN(2).value = 0
    End If
    If (Glo_NoRecOpen4_YN = "Y") Then
        chk_NoRecOpen_YN(3).value = 1
    Else
        chk_NoRecOpen_YN(3).value = 0
    End If
    If (Glo_NoRecOpen5_YN = "Y") Then
        chk_NoRecOpen_YN(4).value = 1
    Else
        chk_NoRecOpen_YN(4).value = 0
    End If
    If (Glo_NoRecOpen6_YN = "Y") Then
        chk_NoRecOpen_YN(5).value = 1
    Else
        chk_NoRecOpen_YN(5).value = 0
    End If

    
'    If Glo_BlackList = "Y" Then
'        chk_BlackList_YN.value = 1
'    Else
'        chk_BlackList_YN.value = 0
'    End If
    '������Ʈ ���
    For i = 0 To MAX_LANE_COUNT - 1
        chk_BlackList_YN(i).Caption = "�̻��": chk_BlackList_YN(i).Enabled = False
    Next i
    If (LANE1_YN = "Y") Then chk_BlackList_YN(0).Caption = LANE1_Name: chk_BlackList_YN(0).Enabled = True
    If (LANE2_YN = "Y") Then chk_BlackList_YN(1).Caption = LANE2_Name: chk_BlackList_YN(1).Enabled = True
    If (LANE3_YN = "Y") Then chk_BlackList_YN(2).Caption = LANE3_Name: chk_BlackList_YN(2).Enabled = True
    If (LANE4_YN = "Y") Then chk_BlackList_YN(3).Caption = LANE4_Name: chk_BlackList_YN(3).Enabled = True
    If (LANE5_YN = "Y") Then chk_BlackList_YN(4).Caption = LANE5_Name: chk_BlackList_YN(4).Enabled = True
    If (LANE6_YN = "Y") Then chk_BlackList_YN(5).Caption = LANE6_Name: chk_BlackList_YN(5).Enabled = True

    If (Glo_BlackList1_YN = "Y") Then
        chk_BlackList_YN(0).value = 1
    Else
        chk_BlackList_YN(0).value = 0
    End If
    If (Glo_BlackList2_YN = "Y") Then
        chk_BlackList_YN(1).value = 1
    Else
        chk_BlackList_YN(1).value = 0
    End If
    If (Glo_BlackList3_YN = "Y") Then
        chk_BlackList_YN(2).value = 1
    Else
        chk_BlackList_YN(2).value = 0
    End If
    If (Glo_BlackList4_YN = "Y") Then
        chk_BlackList_YN(3).value = 1
    Else
        chk_BlackList_YN(3).value = 0
    End If
    If (Glo_BlackList5_YN = "Y") Then
        chk_BlackList_YN(4).value = 1
    Else
        chk_BlackList_YN(4).value = 0
    End If
    If (Glo_BlackList6_YN = "Y") Then
        chk_BlackList_YN(5).value = 1
    Else
        chk_BlackList_YN(5).value = 0
    End If



    '�Ϲ����� �����н� ���
    For i = 0 To MAX_LANE_COUNT - 1
        chk_FreePassLane_YN(i).Caption = "�̻��": chk_FreePassLane_YN(i).Enabled = False
    Next i
    If (LANE1_YN = "Y") Then chk_FreePassLane_YN(0).Caption = LANE1_Name: chk_FreePassLane_YN(0).Enabled = True
    If (LANE2_YN = "Y") Then chk_FreePassLane_YN(1).Caption = LANE2_Name: chk_FreePassLane_YN(1).Enabled = True
    If (LANE3_YN = "Y") Then chk_FreePassLane_YN(2).Caption = LANE3_Name: chk_FreePassLane_YN(2).Enabled = True
    If (LANE4_YN = "Y") Then chk_FreePassLane_YN(3).Caption = LANE4_Name: chk_FreePassLane_YN(3).Enabled = True
    If (LANE5_YN = "Y") Then chk_FreePassLane_YN(4).Caption = LANE5_Name: chk_FreePassLane_YN(4).Enabled = True
    If (LANE6_YN = "Y") Then chk_FreePassLane_YN(5).Caption = LANE6_Name: chk_FreePassLane_YN(5).Enabled = True
    
    If (Glo_FreePassLane1_YN = "Y") Then
        chk_FreePassLane_YN(0).value = 1
    Else
        chk_FreePassLane_YN(0).value = 0
    End If
    If (Glo_FreePassLane2_YN = "Y") Then
        chk_FreePassLane_YN(1).value = 1
    Else
        chk_FreePassLane_YN(1).value = 0
    End If
    If (Glo_FreePassLane3_YN = "Y") Then
        chk_FreePassLane_YN(2).value = 1
    Else
        chk_FreePassLane_YN(2).value = 0
    End If
    If (Glo_FreePassLane4_YN = "Y") Then
        chk_FreePassLane_YN(3).value = 1
    Else
        chk_FreePassLane_YN(3).value = 0
    End If
    If (Glo_FreePassLane5_YN = "Y") Then
        chk_FreePassLane_YN(4).value = 1
    Else
        chk_FreePassLane_YN(4).value = 0
    End If
    If (Glo_FreePassLane6_YN = "Y") Then
        chk_FreePassLane_YN(5).value = 1
    Else
        chk_FreePassLane_YN(5).value = 0
    End If

    
    '�������� �����н� ���
    For i = 0 To MAX_LANE_COUNT - 1
        chk_Taxi_YN(i).Caption = "�̻��": chk_Taxi_YN(i).Enabled = False
    Next i
    If (LANE1_YN = "Y") Then chk_Taxi_YN(0).Caption = LANE1_Name: chk_Taxi_YN(0).Enabled = True
    If (LANE2_YN = "Y") Then chk_Taxi_YN(1).Caption = LANE2_Name: chk_Taxi_YN(1).Enabled = True
    If (LANE3_YN = "Y") Then chk_Taxi_YN(2).Caption = LANE3_Name: chk_Taxi_YN(2).Enabled = True
    If (LANE4_YN = "Y") Then chk_Taxi_YN(3).Caption = LANE4_Name: chk_Taxi_YN(3).Enabled = True
    If (LANE5_YN = "Y") Then chk_Taxi_YN(4).Caption = LANE5_Name: chk_Taxi_YN(4).Enabled = True
    If (LANE6_YN = "Y") Then chk_Taxi_YN(5).Caption = LANE6_Name: chk_Taxi_YN(5).Enabled = True
    
    If (Glo_TAXI1_YN = "Y") Then
        chk_Taxi_YN(0).value = 1
    Else
        chk_Taxi_YN(0).value = 0
    End If
    If (Glo_TAXI2_YN = "Y") Then
        chk_Taxi_YN(1).value = 1
    Else
        chk_Taxi_YN(1).value = 0
    End If
    If (Glo_TAXI3_YN = "Y") Then
        chk_Taxi_YN(2).value = 1
    Else
        chk_Taxi_YN(2).value = 0
    End If
    If (Glo_TAXI4_YN = "Y") Then
        chk_Taxi_YN(3).value = 1
    Else
        chk_Taxi_YN(3).value = 0
    End If
    If (Glo_TAXI5_YN = "Y") Then
        chk_Taxi_YN(4).value = 1
    Else
        chk_Taxi_YN(4).value = 0
    End If
    If (Glo_TAXI6_YN = "Y") Then
        chk_Taxi_YN(5).value = 1
    Else
        chk_Taxi_YN(5).value = 0
    End If
    
    
    
    '�湮���� �湮�� ���� ����(�̵������)
    If (LANE1_YN = "Y") Then
        chk_Guest_YN(0).Enabled = True
        chk_Guest_YN(0).Caption = LANE1_Name
        If (Glo_GUEST_LANE1_YN = "Y") Then chk_Guest_YN(0).value = 1 Else chk_Guest_YN(0).value = 0
    Else
        chk_Guest_YN(0).Enabled = False
        chk_Guest_YN(0).Caption = "�̻��"
    End If
    If (LANE2_YN = "Y") Then
        chk_Guest_YN(1).Enabled = True
        chk_Guest_YN(1).Caption = LANE2_Name
        If (Glo_GUEST_LANE2_YN = "Y") Then chk_Guest_YN(1).value = 1 Else chk_Guest_YN(1).value = 0
    Else
        chk_Guest_YN(1).Enabled = False
        chk_Guest_YN(1).Caption = "�̻��"
    End If
    If (LANE3_YN = "Y") Then
        chk_Guest_YN(2).Enabled = True
        chk_Guest_YN(2).Caption = LANE3_Name
        If (Glo_GUEST_LANE3_YN = "Y") Then chk_Guest_YN(2).value = 1 Else chk_Guest_YN(2).value = 0
    Else
        chk_Guest_YN(2).Enabled = False
        chk_Guest_YN(2).Caption = "�̻��"
    End If
    If (LANE4_YN = "Y") Then
        chk_Guest_YN(3).Enabled = True
        chk_Guest_YN(3).Caption = LANE4_Name
        If (Glo_GUEST_LANE4_YN = "Y") Then chk_Guest_YN(3).value = 1 Else chk_Guest_YN(3).value = 0
    Else
        chk_Guest_YN(3).Enabled = False
        chk_Guest_YN(3).Caption = "�̻��"
    End If
    If (LANE5_YN = "Y") Then
        chk_Guest_YN(4).Enabled = True
        chk_Guest_YN(4).Caption = LANE5_Name
        If (Glo_GUEST_LANE5_YN = "Y") Then chk_Guest_YN(4).value = 1 Else chk_Guest_YN(4).value = 0
    Else
        chk_Guest_YN(4).Enabled = False
        chk_Guest_YN(4).Caption = "�̻��"
    End If
    If (LANE6_YN = "Y") Then
        chk_Guest_YN(5).Enabled = True
        chk_Guest_YN(5).Caption = LANE6_Name
        If (Glo_GUEST_LANE6_YN = "Y") Then chk_Guest_YN(5).value = 1 Else chk_Guest_YN(5).value = 0
    Else
        chk_Guest_YN(5).Enabled = False
        chk_Guest_YN(5).Caption = "�̻��"
    End If
    
    
    
    '�ڸ���� ���� ����
    For i = 0 To MAX_LANE_COUNT - 1
        chk_NoWork_YN(i).Caption = "�̻��": chk_NoWork_YN(i).Enabled = False
    Next i
    If (LANE1_YN = "Y") Then chk_NoWork_YN(0).Caption = LANE1_Name: chk_NoWork_YN(0).Enabled = True
    If (LANE2_YN = "Y") Then chk_NoWork_YN(1).Caption = LANE2_Name: chk_NoWork_YN(1).Enabled = True
    If (LANE3_YN = "Y") Then chk_NoWork_YN(2).Caption = LANE3_Name: chk_NoWork_YN(2).Enabled = True
    If (LANE4_YN = "Y") Then chk_NoWork_YN(3).Caption = LANE4_Name: chk_NoWork_YN(3).Enabled = True
    If (LANE5_YN = "Y") Then chk_NoWork_YN(4).Caption = LANE5_Name: chk_NoWork_YN(4).Enabled = True
    If (LANE6_YN = "Y") Then chk_NoWork_YN(5).Caption = LANE6_Name: chk_NoWork_YN(5).Enabled = True
    If (Glo_NOWORK1_YN = "Y") Then
        chk_NoWork_YN(0).value = 1
    Else
        chk_NoWork_YN(0).value = 0
    End If
    If (Glo_NOWORK2_YN = "Y") Then
        chk_NoWork_YN(1).value = 1
    Else
        chk_NoWork_YN(1).value = 0
    End If
    If (Glo_NOWORK3_YN = "Y") Then
        chk_NoWork_YN(2).value = 1
    Else
        chk_NoWork_YN(2).value = 0
    End If
    If (Glo_NOWORK4_YN = "Y") Then
        chk_NoWork_YN(3).value = 1
    Else
        chk_NoWork_YN(3).value = 0
    End If
    If (Glo_NOWORK5_YN = "Y") Then
        chk_NoWork_YN(4).value = 1
    Else
        chk_NoWork_YN(4).value = 0
    End If
    If (Glo_NOWORK6_YN = "Y") Then
        chk_NoWork_YN(5).value = 1
    Else
        chk_NoWork_YN(5).value = 0
    End If
    
    

    ' ������ ��� �����Ⱓ(�ִ� 36����)
    txt_using_date.text = CStr(Glo_INOUT_USING_DATE)
    
    '�ѱ� ���ν� ���͸� ��뿩��
    If MissMatch_YN = "Y" Then
        chk_MissMatch_YN.value = 1
    Else
        chk_MissMatch_YN.value = 0
    End If

    'Ȩ�� ����
    '���Ϸε�
    iHomeNetNo = Val(Get_Ini2("System Config", "HomeNetMode", "1", "C:\HomeNet\HomeNet.ini"))
    If HomeNet_YN = "Y" Then
        chk_HomeNet_YN.value = 1
        frmHomeNet.Enabled = True
    Else
        chk_HomeNet_YN.value = 0
        frmHomeNet.Enabled = False
    End If
    Text_HomeNet_IP = Trim(HomeNet_IP)
    Text_HomeNet_Port = Val(HomeNet_Port)

    With cmb_HomeNet
        .AddItem "1.�������"
        .AddItem "2.�������(DB)"
        .AddItem "3.������"
        .AddItem "4.����"
        .AddItem "5.�ڸƽ�"
        .AddItem "6.������Ʈ�ѽ�"
        .AddItem "7.�浿 ��"
        .AddItem "8.LG����"
        .AddItem "9.�������(TCP)"
        .AddItem "10.�������(����������)"
        .AddItem "11.�Ƽ�����(GS �׿���)"
        .AddItem "12.ȨŬ����"
    End With
    cmb_HomeNet.text = cmb_HomeNet.List(iHomeNetNo - 1)


    If (MissMatch_HomeNet_YN = "Y") Then
        chk_MissMatch_HomeNet_YN.value = 1
    Else
        chk_MissMatch_HomeNet_YN.value = 0
    End If
    
    
    

    If MVR_YN = "Y" Then
        chk_MVR_YN.value = 1
        frmMVR.Enabled = True
    Else
        chk_MVR_YN.value = 0
        frmMVR.Enabled = False
    End If
    Text_MVR_IP = Trim(MVR_IP)
    Text_MVR_Port = Val(MVR_Port)
    

    '������
    If Glo_WEEK_YN = "Y" Then
        chk_Week_YN.value = 1
    Else
        chk_Week_YN.value = 0
    End If
    

    '��������
    With cmb_Rotation
        .AddItem "������"
        .AddItem "2����"
        .AddItem "5����"
        .AddItem "10����"
    End With
    cmb_Rotation.text = Glo_ROTATION
'    If (Glo_ROTATION = "������") Then
'        cmb_Rotation.Index = 0
'    ElseIf (Glo_ROTATION = "2����") Then
'        cmb_Rotation.Index = 1
'    ElseIf (Glo_ROTATION = "5����") Then
'        cmb_Rotation.Index = 2
'    ElseIf (Glo_ROTATION = "10����") Then
'        cmb_Rotation.Index = 3
'    Else
'        cmb_Rotation.Index = 0
'    End If

    
    '��������� ������¥ �ʱ�ȭ
    DTPicker1.value = Format(Now, "yyyy-mm-dd")
    
    
    '������ ����
    Dim rs As Recordset
    Dim sQry As String
    Dim bQryResult As Boolean
    
    sQry = "Select * From tb_config WHERE NAME = 'WebDC' "
    Set rs = New ADODB.Recordset
    rs.Open sQry, adoConn
    If (rs.EOF) Then
    Else
        If (rs!Content = "Y") Then
            chk_WebDC.value = 1
        Else
            chk_WebDC.value = 0
        End If
    End If
    Set rs = Nothing
    '������ ��
    
    
    '�����湮������� ����
'''    sQry = "Select * From tb_config WHERE NAME = 'GuestCarReg' "
'''    Set rs = New ADODB.Recordset
'''    rs.Open sQry, adoConn
'''    If (rs.EOF) Then
'''    Else
'''        If (rs!Content = "Y") Then
'''            chk_GuestCarReg.value = 1
'''        Else
'''            chk_GuestCarReg.value = 0
'''        End If
'''    End If
'''    Set rs = Nothing
    
    If (Glo_GuestReg_YN = "Y") Then
        chk_GuestCarReg.value = 1
    Else
        chk_GuestCarReg.value = 0
    End If
    '�����湮������� ��
    
    '����� �˸� ��� ����
    If (Glo_MobileAlarm_YN = "Y") Then
        chk_MobileAlarm.value = 1
    Else
        chk_MobileAlarm.value = 0
    End If
    '����� �˸� ��� ��
    
    
    
    Exit Sub
    
Err_p:
    DataLogger ("FrmExtend : " & Err.Description)

End Sub

Private Sub Text_HomeNet_IP_Change()
    If (HomeNet_IP <> Text_HomeNet_IP.text) Then
        cmd_HomeTest.Enabled = False
    Else
        cmd_HomeTest.Enabled = True
    End If
End Sub
