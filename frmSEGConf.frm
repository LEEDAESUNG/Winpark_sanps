VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmSEGConf 
   BorderStyle     =   1  '단일 고정
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11895
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSEGConf.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   527
   ScaleMode       =   3  '픽셀
   ScaleWidth      =   793
   StartUpPosition =   2  '화면 가운데
   Begin VB.Frame Frame12 
      Caption         =   "PPPoE Information"
      Height          =   1215
      Left            =   120
      TabIndex        =   122
      Top             =   5760
      Width           =   2895
      Begin VB.TextBox txtPass 
         BackColor       =   &H80000011&
         Enabled         =   0   'False
         Height          =   330
         IMEMode         =   3  '사용 못함
         Left            =   1080
         PasswordChar    =   "*"
         TabIndex        =   124
         Top             =   720
         Width           =   1680
      End
      Begin VB.TextBox txtID 
         BackColor       =   &H80000011&
         Enabled         =   0   'False
         Height          =   330
         Left            =   1080
         TabIndex        =   123
         Top             =   360
         Width           =   1680
      End
      Begin VB.Label Label26 
         Caption         =   "Password"
         Height          =   255
         Left            =   120
         TabIndex        =   126
         Top             =   765
         Width           =   975
      End
      Begin VB.Label Label24 
         Caption         =   "PPPoE ID"
         Height          =   255
         Left            =   120
         TabIndex        =   125
         Top             =   405
         Width           =   855
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "IP Address Information"
      Height          =   1575
      Left            =   120
      TabIndex        =   115
      Top             =   4080
      Width           =   2895
      Begin VB.TextBox txtGW 
         Height          =   285
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   118
         Top             =   1080
         Width           =   1680
      End
      Begin VB.TextBox txtSubnet 
         Height          =   330
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   117
         Top             =   690
         Width           =   1680
      End
      Begin VB.TextBox txtIP 
         Height          =   285
         Left            =   1110
         MaxLength       =   15
         TabIndex        =   116
         Top             =   360
         Width           =   1680
      End
      Begin VB.Label Label4 
         Caption         =   "Gateway "
         Height          =   255
         Left            =   240
         TabIndex        =   121
         Top             =   1095
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Subnet "
         Height          =   255
         Left            =   360
         TabIndex        =   120
         Top             =   735
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Local IP "
         Height          =   255
         Left            =   120
         TabIndex        =   119
         Top             =   375
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "IP Configuration Method"
      Height          =   735
      Left            =   120
      TabIndex        =   111
      Top             =   3240
      Width           =   2895
      Begin VB.OptionButton Option3 
         Caption         =   "PPPoE"
         Height          =   225
         Left            =   1920
         TabIndex        =   114
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         Caption         =   "DHCP"
         Height          =   225
         Left            =   1080
         TabIndex        =   113
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Static"
         Height          =   225
         Left            =   120
         TabIndex        =   112
         Top             =   360
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.TextBox txtConnect1 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   330
      Left            =   9840
      MaxLength       =   15
      TabIndex        =   88
      Top             =   120
      Width           =   1800
   End
   Begin VB.TextBox txtConnect0 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   330
      Left            =   7200
      MaxLength       =   15
      TabIndex        =   9
      Top             =   120
      Width           =   1800
   End
   Begin VB.Frame FrameTab 
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5970
      Index           =   2
      Left            =   9720
      TabIndex        =   6
      Top             =   8160
      Visible         =   0   'False
      Width           =   8400
      Begin VB.Frame Frame10 
         Height          =   1335
         Left            =   3720
         TabIndex        =   99
         Top             =   0
         Width           =   4695
         Begin VB.TextBox txtITime1 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   270
            Left            =   2040
            MaxLength       =   15
            TabIndex        =   100
            Top             =   360
            Width           =   840
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "(0 ~ 65535 sec)"
            Height          =   225
            Index           =   1
            Left            =   3000
            TabIndex        =   103
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label21 
            Caption         =   "* Closes socket connection, if there is      no transmission during this time."
            Height          =   585
            Index           =   1
            Left            =   120
            TabIndex        =   102
            Top             =   720
            Width           =   4455
         End
         Begin VB.Label Label12 
            Caption         =   "Inactivity time"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   101
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Destination Information"
         Height          =   1335
         Left            =   120
         TabIndex        =   81
         Top             =   4560
         Width           =   8295
         Begin VB.CheckBox ChkDNS1 
            Appearance      =   0  '평면
            Caption         =   "Use DNS"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   6960
            TabIndex        =   109
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txt_DNS_ServerIP1 
            Enabled         =   0   'False
            Height          =   330
            Left            =   4560
            TabIndex        =   108
            Top             =   345
            Width           =   2055
         End
         Begin VB.TextBox txtServerIP1 
            Height          =   330
            Left            =   1200
            MaxLength       =   15
            TabIndex        =   106
            Top             =   315
            Width           =   1665
         End
         Begin VB.TextBox txtServer_Domain1 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4560
            TabIndex        =   104
            Top             =   840
            Width           =   3495
         End
         Begin VB.TextBox txtServerPort1 
            Height          =   330
            Left            =   1200
            MaxLength       =   15
            TabIndex        =   82
            Top             =   810
            Width           =   960
         End
         Begin VB.Label lDNSServerIP1 
            Caption         =   "DNS Server IP"
            Enabled         =   0   'False
            Height          =   255
            Left            =   3000
            TabIndex        =   110
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label lPeerIP1 
            Caption         =   "Peer IP"
            Height          =   255
            Left            =   120
            TabIndex        =   107
            Top             =   360
            Width           =   1200
         End
         Begin VB.Label lDomainName1 
            Caption         =   "Domain Name"
            Enabled         =   0   'False
            Height          =   255
            Left            =   3240
            TabIndex        =   105
            Top             =   855
            Width           =   1335
         End
         Begin VB.Label lPeerPort1 
            Caption         =   "Peer Port"
            Height          =   255
            Left            =   120
            TabIndex        =   83
            Top             =   855
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Data Packing Condition"
         Height          =   1335
         Index           =   1
         Left            =   3720
         TabIndex        =   60
         Top             =   1440
         Width           =   4695
         Begin VB.TextBox txtDChar1 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   270
            Left            =   1320
            MaxLength       =   2
            TabIndex        =   63
            Top             =   960
            Width           =   840
         End
         Begin VB.TextBox txtDSize1 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   270
            Left            =   1320
            MaxLength       =   15
            TabIndex        =   62
            Top             =   600
            Width           =   840
         End
         Begin VB.TextBox txtDTime1 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   270
            Left            =   1320
            MaxLength       =   15
            TabIndex        =   61
            Top             =   240
            Width           =   840
         End
         Begin VB.Label Label11 
            Caption         =   "(0 ~ 65535 ms)"
            Height          =   255
            Index           =   1
            Left            =   2280
            TabIndex        =   69
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label20 
            Caption         =   "(Hexacode)"
            Height          =   255
            Index           =   1
            Left            =   2280
            TabIndex        =   68
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label23 
            Caption         =   "(0 ~ 255 Byte)"
            Height          =   255
            Index           =   1
            Left            =   2280
            TabIndex        =   67
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label19 
            Caption         =   "Char"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   66
            Top             =   960
            Width           =   495
         End
         Begin VB.Label Label17 
            Caption         =   "Size"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   65
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label16 
            Caption         =   "Time"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   64
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Password (TCP Server)"
         Height          =   735
         Index           =   1
         Left            =   3720
         TabIndex        =   56
         Top             =   3720
         Width           =   4695
         Begin VB.TextBox txtTCPPass1 
            Height          =   270
            Left            =   2520
            TabIndex        =   58
            Top             =   345
            Width           =   975
         End
         Begin VB.CheckBox chkTCPPass1 
            Caption         =   "Enable"
            Height          =   225
            Left            =   240
            TabIndex        =   57
            Top             =   375
            Width           =   975
         End
         Begin VB.Label Label27 
            Caption         =   "Password"
            Height          =   255
            Index           =   1
            Left            =   1560
            TabIndex        =   59
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Serial 1 Setup"
         Height          =   2775
         Index           =   1
         Left            =   120
         TabIndex        =   45
         Top             =   0
         Width           =   3495
         Begin VB.ComboBox cboDataBits1 
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1560
            Style           =   2  '드롭다운 목록
            TabIndex        =   50
            Top             =   840
            Width           =   1740
         End
         Begin VB.ComboBox cboParity1 
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1560
            Style           =   2  '드롭다운 목록
            TabIndex        =   49
            Top             =   1320
            Width           =   1740
         End
         Begin VB.ComboBox cboStopBits1 
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1560
            Style           =   2  '드롭다운 목록
            TabIndex        =   48
            Top             =   1800
            Width           =   1740
         End
         Begin VB.ComboBox cboSpeed1 
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "frmSEGConf.frx":0E42
            Left            =   1560
            List            =   "frmSEGConf.frx":0E44
            Style           =   2  '드롭다운 목록
            TabIndex        =   47
            Top             =   360
            Width           =   1740
         End
         Begin VB.ComboBox cboFlow1 
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1560
            Style           =   2  '드롭다운 목록
            TabIndex        =   46
            Top             =   2280
            Width           =   1740
         End
         Begin VB.Label Label9 
            Alignment       =   1  '오른쪽 맞춤
            Caption         =   "DataBit"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   360
            TabIndex        =   55
            Top             =   930
            Width           =   915
         End
         Begin VB.Label Label8 
            Alignment       =   1  '오른쪽 맞춤
            Caption         =   "Parity"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   240
            TabIndex        =   54
            Top             =   1380
            Width           =   1005
         End
         Begin VB.Label Label7 
            Alignment       =   1  '오른쪽 맞춤
            Caption         =   "Stop Bit"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   240
            TabIndex        =   53
            Top             =   1830
            Width           =   1005
         End
         Begin VB.Label Label6 
            Alignment       =   1  '오른쪽 맞춤
            Caption         =   "Speed"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   240
            TabIndex        =   52
            Top             =   480
            Width           =   1005
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            Caption         =   "Flow"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   720
            TabIndex        =   51
            Top             =   2280
            Width           =   525
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Operation Mode"
         Height          =   1575
         Index           =   1
         Left            =   120
         TabIndex        =   70
         Top             =   2880
         Width           =   3495
         Begin VB.TextBox txtPort1 
            Height          =   285
            Left            =   2520
            MaxLength       =   15
            TabIndex        =   86
            Top             =   480
            Width           =   840
         End
         Begin VB.CheckBox chkUDPMode1 
            Appearance      =   0  '평면
            Caption         =   "Use UDP mode"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1320
            TabIndex        =   75
            Top             =   960
            Width           =   1575
         End
         Begin VB.OptionButton optClientMode1 
            Caption         =   "Client"
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   73
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton optClientMode1 
            Caption         =   "Mixed"
            Height          =   225
            Index           =   1
            Left            =   120
            TabIndex        =   72
            Top             =   1200
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optClientMode1 
            Caption         =   "Server"
            Height          =   225
            Index           =   2
            Left            =   120
            TabIndex        =   71
            Top             =   780
            Width           =   975
         End
         Begin VB.Label Label30 
            Caption         =   "Local Port"
            Height          =   255
            Left            =   1320
            TabIndex        =   87
            Top             =   480
            Width           =   1215
         End
      End
   End
   Begin VB.Frame FrameTab 
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5970
      Index           =   1
      Left            =   900
      TabIndex        =   5
      Top             =   8160
      Visible         =   0   'False
      Width           =   8400
      Begin VB.Frame Frame9 
         Height          =   1335
         Left            =   3720
         TabIndex        =   94
         Top             =   0
         Width           =   4695
         Begin VB.TextBox txtITime0 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   270
            Left            =   2040
            MaxLength       =   15
            TabIndex        =   95
            Top             =   360
            Width           =   840
         End
         Begin VB.Label Label12 
            Caption         =   "Inactivity time"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   98
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label21 
            Caption         =   "* Closes socket connection, if there is      no transmission during this time."
            Height          =   585
            Index           =   0
            Left            =   120
            TabIndex        =   97
            Top             =   720
            Width           =   4455
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "(0 ~ 65535 sec)"
            Height          =   225
            Index           =   0
            Left            =   3000
            TabIndex        =   96
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Destination Information"
         Height          =   1335
         Left            =   120
         TabIndex        =   76
         Top             =   4560
         Width           =   8295
         Begin VB.TextBox txt_DNS_ServerIP0 
            Enabled         =   0   'False
            Height          =   330
            Left            =   4560
            TabIndex        =   92
            Top             =   345
            Width           =   2055
         End
         Begin VB.TextBox txtServer_Domain0 
            Enabled         =   0   'False
            Height          =   285
            Left            =   4560
            TabIndex        =   90
            Top             =   840
            Width           =   3495
         End
         Begin VB.CheckBox ChkDNS0 
            Appearance      =   0  '평면
            Caption         =   "Use DNS"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   6960
            TabIndex        =   89
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtServerPort0 
            Height          =   330
            Left            =   1200
            MaxLength       =   15
            TabIndex        =   78
            Top             =   810
            Width           =   960
         End
         Begin VB.TextBox txtServerIP0 
            Height          =   330
            Left            =   1200
            MaxLength       =   15
            TabIndex        =   77
            Top             =   315
            Width           =   1665
         End
         Begin VB.Label lDNSServerIP0 
            Caption         =   "DNS Server IP"
            Enabled         =   0   'False
            Height          =   255
            Left            =   3000
            TabIndex        =   93
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label lDomainName0 
            Caption         =   "Domain Name"
            Enabled         =   0   'False
            Height          =   255
            Left            =   3240
            TabIndex        =   91
            Top             =   855
            Width           =   1335
         End
         Begin VB.Label lPeerPort0 
            Caption         =   "Peer Port"
            Height          =   255
            Left            =   120
            TabIndex        =   80
            Top             =   855
            Width           =   1215
         End
         Begin VB.Label lPeerIP0 
            Caption         =   "Peer IP"
            Height          =   255
            Left            =   120
            TabIndex        =   79
            Top             =   360
            Width           =   1200
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Serial 0 Setup"
         Height          =   2775
         Index           =   0
         Left            =   120
         TabIndex        =   34
         Top             =   0
         Width           =   3495
         Begin VB.ComboBox cboFlow0 
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1560
            Style           =   2  '드롭다운 목록
            TabIndex        =   39
            Top             =   2280
            Width           =   1740
         End
         Begin VB.ComboBox cboSpeed0 
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "frmSEGConf.frx":0E46
            Left            =   1560
            List            =   "frmSEGConf.frx":0E48
            Style           =   2  '드롭다운 목록
            TabIndex        =   38
            Top             =   360
            Width           =   1740
         End
         Begin VB.ComboBox cboStopBits0 
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1560
            Style           =   2  '드롭다운 목록
            TabIndex        =   37
            Top             =   1800
            Width           =   1740
         End
         Begin VB.ComboBox cboParity0 
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1560
            Style           =   2  '드롭다운 목록
            TabIndex        =   36
            Top             =   1320
            Width           =   1740
         End
         Begin VB.ComboBox cboDataBits0 
            Appearance      =   0  '평면
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1560
            Style           =   2  '드롭다운 목록
            TabIndex        =   35
            Top             =   840
            Width           =   1740
         End
         Begin VB.Label Label2 
            Alignment       =   1  '오른쪽 맞춤
            Caption         =   "Flow"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   0
            Left            =   720
            TabIndex        =   44
            Top             =   2280
            Width           =   525
         End
         Begin VB.Label Label6 
            Alignment       =   1  '오른쪽 맞춤
            Caption         =   "Speed"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   240
            TabIndex        =   43
            Top             =   480
            Width           =   1005
         End
         Begin VB.Label Label7 
            Alignment       =   1  '오른쪽 맞춤
            Caption         =   "Stop Bit"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   240
            TabIndex        =   42
            Top             =   1830
            Width           =   1005
         End
         Begin VB.Label Label8 
            Alignment       =   1  '오른쪽 맞춤
            Caption         =   "Parity"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   240
            TabIndex        =   41
            Top             =   1380
            Width           =   1005
         End
         Begin VB.Label Label9 
            Alignment       =   1  '오른쪽 맞춤
            Caption         =   "DataBit"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   360
            TabIndex        =   40
            Top             =   930
            Width           =   915
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Serial Configuration"
         Height          =   735
         Index           =   0
         Left            =   3720
         TabIndex        =   28
         Top             =   2880
         Width           =   4695
         Begin VB.CheckBox CheckSCfg0 
            Caption         =   "Enable"
            Height          =   255
            Left            =   240
            TabIndex        =   32
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox TextSCfg0 
            Height          =   240
            Index           =   0
            Left            =   3120
            TabIndex        =   31
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox TextSCfg0 
            Height          =   240
            Index           =   1
            Left            =   3600
            TabIndex        =   30
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox TextSCfg0 
            Height          =   240
            Index           =   2
            Left            =   4080
            TabIndex        =   29
            Top             =   360
            Width           =   375
         End
         Begin VB.Label Label22 
            Caption         =   "String(as hex)"
            Height          =   255
            Index           =   0
            Left            =   1560
            TabIndex        =   33
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Password (TCP Server)"
         Height          =   735
         Index           =   0
         Left            =   3720
         TabIndex        =   24
         Top             =   3720
         Width           =   4695
         Begin VB.CheckBox chkTCPPass0 
            Caption         =   "Enable"
            Height          =   225
            Left            =   240
            TabIndex        =   26
            Top             =   375
            Width           =   975
         End
         Begin VB.TextBox txtTCPPass0 
            Height          =   270
            Left            =   2520
            TabIndex        =   25
            Top             =   345
            Width           =   975
         End
         Begin VB.Label Label27 
            Caption         =   "Password"
            Height          =   255
            Index           =   0
            Left            =   1560
            TabIndex        =   27
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Data Packing Condition"
         Height          =   1335
         Index           =   0
         Left            =   3720
         TabIndex        =   14
         Top             =   1440
         Width           =   4695
         Begin VB.TextBox txtDTime0 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   270
            Left            =   1320
            MaxLength       =   15
            TabIndex        =   17
            Top             =   240
            Width           =   840
         End
         Begin VB.TextBox txtDSize0 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   270
            Left            =   1320
            MaxLength       =   15
            TabIndex        =   16
            Top             =   600
            Width           =   840
         End
         Begin VB.TextBox txtDChar0 
            Alignment       =   1  '오른쪽 맞춤
            Height          =   270
            Left            =   1320
            MaxLength       =   2
            TabIndex        =   15
            Top             =   960
            Width           =   840
         End
         Begin VB.Label Label16 
            Caption         =   "Time"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   720
            TabIndex        =   23
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label17 
            Caption         =   "Size"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   720
            TabIndex        =   22
            Top             =   600
            Width           =   495
         End
         Begin VB.Label Label19 
            Caption         =   "Char"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   720
            TabIndex        =   21
            Top             =   960
            Width           =   495
         End
         Begin VB.Label Label23 
            Caption         =   "(0 ~ 255 Byte)"
            Height          =   255
            Index           =   0
            Left            =   2280
            TabIndex        =   20
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label20 
            Caption         =   "(Hexacode)"
            Height          =   255
            Index           =   0
            Left            =   2280
            TabIndex        =   19
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label11 
            Caption         =   "(0 ~ 65535 ms)"
            Height          =   255
            Index           =   0
            Left            =   2280
            TabIndex        =   18
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Operation Mode"
         Height          =   1575
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   2880
         Width           =   3495
         Begin VB.TextBox txtPort0 
            Height          =   285
            Left            =   2520
            MaxLength       =   15
            TabIndex        =   84
            Top             =   480
            Width           =   840
         End
         Begin VB.CheckBox chkUDPMode0 
            Appearance      =   0  '평면
            Caption         =   "Use UDP mode"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1320
            TabIndex        =   74
            Top             =   960
            Width           =   1575
         End
         Begin VB.OptionButton optClientMode0 
            Caption         =   "Server"
            Height          =   225
            Index           =   2
            Left            =   120
            TabIndex        =   13
            Top             =   780
            Width           =   975
         End
         Begin VB.OptionButton optClientMode0 
            Caption         =   "Mixed"
            Height          =   225
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   1200
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optClientMode0 
            Caption         =   "Client"
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label13 
            Caption         =   "Local Port"
            Height          =   255
            Left            =   1320
            TabIndex        =   85
            Top             =   480
            Width           =   1215
         End
      End
   End
   Begin VB.CheckBox chkDirect 
      Appearance      =   0  '평면
      Caption         =   "Direct IP Search"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   7365
      Width           =   2415
   End
   Begin VB.TextBox txtDirectIP 
      Height          =   360
      Left            =   6120
      MaxLength       =   15
      TabIndex        =   7
      Top             =   7320
      Width           =   2475
   End
   Begin MSWinsockLib.Winsock WinsockDirect 
      Left            =   12240
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlListView 
      Left            =   14160
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSEGConf.frx":0E4A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog OpenLog 
      Left            =   14880
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox chkDebug 
      Appearance      =   0  '평면
      Caption         =   "Enable Serial Debug Mode"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
   Begin VB.TextBox txtVersion 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   330
      Left            =   1440
      MaxLength       =   15
      TabIndex        =   0
      Top             =   120
      Width           =   1080
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   13440
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSEGConf.frx":1724
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSEGConf.frx":1FFE
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSEGConf.frx":28D8
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSEGConf.frx":31B2
            Key             =   "IMG4"
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock WinsockUDP 
      Left            =   12840
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   765
      Left            =   9000
      TabIndex        =   3
      Top             =   7080
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1349
      ButtonWidth     =   1244
      ButtonHeight    =   1349
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
            Key             =   "SearchBoard"
            Object.ToolTipText     =   "Search Board"
            ImageKey        =   "IMG1"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Setting"
            Key             =   "SettingBoard"
            Object.ToolTipText     =   "Setting Board Information"
            ImageKey        =   "IMG2"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Upload"
            Key             =   "Upload"
            Object.ToolTipText     =   "Upload Firmware"
            ImageKey        =   "IMG3"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "Exit"
            Object.ToolTipText     =   "Exit"
            ImageKey        =   "IMG4"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   6420
      Left            =   3120
      TabIndex        =   4
      Top             =   600
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   11324
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "UART 0"
            Key             =   "UART0"
            Object.Tag             =   "UART0"
            Object.ToolTipText     =   "UART 0 Configuration"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "UART 1"
            Key             =   "UART 1"
            Object.Tag             =   "UART 1"
            Object.ToolTipText     =   "UART 1 Configuration"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListBoards 
      Height          =   2535
      Left            =   120
      TabIndex        =   127
      Top             =   600
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imlListView"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label10 
      Caption         =   "PORT 1"
      Height          =   255
      Index           =   1
      Left            =   9120
      TabIndex        =   129
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label10 
      Caption         =   "PORT 0"
      Height          =   255
      Index           =   0
      Left            =   6480
      TabIndex        =   128
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label5 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "Version"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmSEGConf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------
' WIZSEGConf - WIZ120SR Configuration Tool
'
' Copyright (c) 2008, WIZnet Inc.
'--------------------------------------------------
Option Explicit

Private mintCurFrame As Integer ' Current TAB Frame

Private Sub ChkDNS0_Click()
  If ChkDNS0.value = 1 Then
    txt_DNS_ServerIP0.Enabled = True
    txtServer_Domain0.Enabled = True
    lDNSServerIP0.Enabled = True
    lDomainName0.Enabled = True
    lPeerIP0.Enabled = False
  End If
  
  If ChkDNS0.value = 0 Then
    txt_DNS_ServerIP0.Enabled = False
    txtServer_Domain0.Enabled = False
    lDNSServerIP0.Enabled = False
    lDomainName0.Enabled = False
    lPeerIP0.Enabled = True
  End If

End Sub
Private Sub ChkDNS1_Click()
  If ChkDNS1.value = 1 Then
    txt_DNS_ServerIP1.Enabled = True
    txtServer_Domain1.Enabled = True
    lDNSServerIP1.Enabled = True
    lDomainName1.Enabled = True
    lPeerIP1.Enabled = False
  End If
  
  If ChkDNS1.value = 0 Then
    txt_DNS_ServerIP1.Enabled = False
    txtServer_Domain1.Enabled = False
    lDNSServerIP1.Enabled = False
    lDomainName1.Enabled = False
    lPeerIP1.Enabled = True
  End If

End Sub

Private Sub chkUDPMode0_Click()
    If chkUDPMode0.value = 1 Then
        optClientMode0(0).Enabled = False
        optClientMode0(1).Enabled = False
        optClientMode0(2).Enabled = False
    Else
        optClientMode0(0).Enabled = True
        optClientMode0(1).Enabled = True
        optClientMode0(2).Enabled = True
    End If
    
End Sub

Private Sub chkUDPMode1_Click()
    If chkUDPMode1.value = 1 Then
        optClientMode1(0).Enabled = False
        optClientMode1(1).Enabled = False
        optClientMode1(2).Enabled = False
    Else
        optClientMode1(0).Enabled = True
        optClientMode1(1).Enabled = True
        optClientMode1(2).Enabled = True
    End If
End Sub

Private Sub Option1_Click()
    
    If Option1.value = True Then
        txtIP.Enabled = True
        txtIP.BackColor = &H80000005
        txtPort0.Enabled = True
        txtPort0.BackColor = &H80000005
        txtSubnet.Enabled = True
        txtSubnet.BackColor = &H80000005
        txtGW.Enabled = True
        txtGW.BackColor = &H80000005
    
        txtID.Enabled = False
        txtID.BackColor = &H80000011
        txtPass.Enabled = False
        txtPass.BackColor = &H80000011
    End If

End Sub

Private Sub Option2_Click()

    If Option2.value = True Then
        txtIP.Enabled = False
        txtIP.BackColor = &H80000011
        txtPort0.Enabled = False
        txtPort0.BackColor = &H80000011
        txtSubnet.Enabled = False
        txtSubnet.BackColor = &H80000011
        txtGW.Enabled = False
        txtGW.BackColor = &H80000011
    
        txtID.Enabled = False
        txtID.BackColor = &H80000011
        txtPass.Enabled = False
        txtPass.BackColor = &H80000011
    End If
    
End Sub

Private Sub Option3_Click()
    
    If Option3.value = True Then
        txtIP.Enabled = False
        txtIP.BackColor = &H80000011
        txtPort0.Enabled = False
        txtPort0.BackColor = &H80000011
        txtSubnet.Enabled = False
        txtSubnet.BackColor = &H80000011
        txtGW.Enabled = False
        txtGW.BackColor = &H80000011
        
        txtID.Enabled = True
        txtID.BackColor = &H80000005
        txtPass.Enabled = True
        txtPass.BackColor = &H80000005
    End If
    
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''
' TAB select
'
''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub TabStrip1_Click()

   If TabStrip1.SelectedItem.index = mintCurFrame Then Exit Sub
   
   ' View Selected Frame
   FrameTab(TabStrip1.SelectedItem.index).Visible = True
   FrameTab(mintCurFrame).Visible = False
   mintCurFrame = TabStrip1.SelectedItem.index

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''
' Name : ShowMsgWindow
' Parameter : none
'
' Show the message for process the action.
'
''''''''''''''''''''''''''''''''''''''''''''''''
Sub ShowMsgWindow()
    
    Me.Enabled = False
    
    frmState.ShowMsg
    frmState.Show vbModal, Me
    
    Me.Enabled = True
    
End Sub
        

''''''''''''''''''''''''''''''''''''''''''''''''
' Name : BoardAdd
' Parameter : newStr() is string of board's configuration data
'
' Add Board with receiving new board's configuration data
' save total count of boards to "intBoardNum" variable.
' save configuration data to "colBoards" collection
'
''''''''''''''''''''''''''''''''''''''''''''''''
Sub BoardAdd(newStr() As Byte)
On Error GoTo e_go

    Dim mac As String
    Dim i As Integer
    
    ' making mac address string key
    ' ex) 00:44:34:EA:3A:F0
    mac = ""
    For i = 0 To 5
        If Len(Hex(newStr(i))) = 1 Then
            mac = mac & "0" & Hex(newStr(i)) & ":"
        Else
            mac = mac & Hex(newStr(i)) & ":"
        End If
    Next i
    mac = Left(mac, Len(mac) - 1)
    
    ' Add Board entity by using mac .
    colBoards.Add newStr, mac
    ' add list view
    frmSEGConf.ListBoards.ListItems.Add intBoardNum, mac, mac
    frmSEGConf.ListBoards.ListItems.Item(intBoardNum).SmallIcon = 1
            
    ' Automatically select the first row of ListView
    If intBoardNum = 1 Then
        Call frmSEGConf.ListBoards_FirstRowSelect
        'frmSEGConf.ListBoards.SetFocus
    End If
    
    intBoardNum = intBoardNum + 1

e_go:
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''
' Name : BoardUpdate
' Parameter : newStr() is string of board's configuration data
'
' Update Board's configuration data with receiving data
' delete previous data
' add new board's data
'
''''''''''''''''''''''''''''''''''''''''''''''''
Sub BoardUpdate(newStr() As Byte)
On Error GoTo U_ERROR

    Dim newInfo() As Byte
    ReDim newInfo(0 To Len(BoardInfo) - 1) As Byte
    
    ' Verify message
    If newStr(0) = BoardInfo.mac(0) And _
        newStr(1) = BoardInfo.mac(1) And _
        newStr(2) = BoardInfo.mac(2) And _
        newStr(3) = BoardInfo.mac(3) And _
        newStr(4) = BoardInfo.mac(4) And _
        newStr(5) = BoardInfo.mac(5) Then
        
        ' Updatae the item
        colBoards.Remove BoardKey
        CopyMemory newInfo(0), BoardInfo, Len(BoardInfo)
        colBoards.Add newInfo, BoardKey
        ToolMode = modeSettingComplete
        
    End If

U_ERROR:
    Erase newInfo
    
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''
' Name : BoardRemove
' Parameter : None
'
' Delete all Board's data from "colBoards" collection
' Board's ListBox clears
' "intBoardNum" variable sets '0'
'
''''''''''''''''''''''''''''''''''''''''''''''''
Sub BoardRemove()
    
    Dim num As Integer
    
    ' Set false the flag, board select
    bSelect = False
    
    ' Delete All Board information
    If intBoardNum > 1 Then
    For num = 1 To intBoardNum - 1
        colBoards.Remove frmSEGConf.ListBoards.ListItems(num).KEY
    Next num
    End If
    
    frmSEGConf.ListBoards.ListItems.Clear
    
    'Clear board's count
    intBoardNum = 1
    
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''
' Name : Form_Load
' Parameter : None
'
' Initialize control, variable, position
'
''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()

    frmSEGConf.Caption = "WIZ120SR Configuration Tool ver " & App.Major & "." & App.Minor & "." & App.Revision
    
    Dim colX As ColumnHeader
    Dim intX As Integer
    'Set colX = ListBoards.ColumnHeaders.Add()
    'colX.Text = "Board list"
    'colX.Width = ListBoards.Width
    ListBoards.ColumnHeaders.Add , , "Board list"


    WinsockUDP.RemoteHost = "255.255.255.255"
    WinsockUDP.RemotePort = 1460
'    WinsockUDP.LocalPort = 5005
    WinsockUDP.Bind
    
    'Speed Value
    cboSpeed0.AddItem "1200", 0
    cboSpeed0.ItemData(0) = &HA0
    cboSpeed0.AddItem "2400", 1
    cboSpeed0.ItemData(1) = &HD0
    cboSpeed0.AddItem "4800", 2
    cboSpeed0.ItemData(2) = &HE8
    cboSpeed0.AddItem "9600", 3
    cboSpeed0.ItemData(3) = &HF4
    cboSpeed0.AddItem "19200", 4
    cboSpeed0.ItemData(4) = &HFA
    cboSpeed0.AddItem "38400", 5
    cboSpeed0.ItemData(5) = &HFD
    cboSpeed0.AddItem "57600", 6
    cboSpeed0.ItemData(6) = &HFE
    cboSpeed0.AddItem "115200", 7
    cboSpeed0.ItemData(7) = &HFF
    cboSpeed0.AddItem "230400", 8
    cboSpeed0.ItemData(8) = &HBB
    
    cboSpeed1.AddItem "1200", 0
    cboSpeed1.ItemData(0) = &HA0
    cboSpeed1.AddItem "2400", 1
    cboSpeed1.ItemData(1) = &HD0
    cboSpeed1.AddItem "4800", 2
    cboSpeed1.ItemData(2) = &HE8
    cboSpeed1.AddItem "9600", 3
    cboSpeed1.ItemData(3) = &HF4
    cboSpeed1.AddItem "19200", 4
    cboSpeed1.ItemData(4) = &HFA
    cboSpeed1.AddItem "38400", 5
    cboSpeed1.ItemData(5) = &HFD
    cboSpeed1.AddItem "57600", 6
    cboSpeed1.ItemData(6) = &HFE
    cboSpeed1.AddItem "115200", 7
    cboSpeed1.ItemData(7) = &HFF
    cboSpeed1.AddItem "230400", 8
    cboSpeed1.ItemData(8) = &HBB
    
    'Databit Value
    cboDataBits0.AddItem "7", 0
    cboDataBits0.ItemData(0) = &H7
    cboDataBits0.AddItem "8", 1
    cboDataBits0.ItemData(1) = &H8
    
    cboDataBits1.AddItem "7", 0
    cboDataBits1.ItemData(0) = &H7
    cboDataBits1.AddItem "8", 1
    cboDataBits1.ItemData(1) = &H8
    
    'Stopbit Value
    cboStopBits0.AddItem "1", 0
    cboStopBits0.ItemData(0) = &H1
    cboStopBits0.AddItem "2", 1
    cboStopBits0.ItemData(1) = &H2
    
    cboStopBits1.AddItem "1", 0
    cboStopBits1.ItemData(0) = &H1
    cboStopBits1.AddItem "2", 1
    cboStopBits1.ItemData(1) = &H2
       
    'Parity Value
    cboParity0.AddItem "None", 0
    cboParity0.ItemData(0) = &H0
    cboParity0.AddItem "Odd", 1
    cboParity0.ItemData(1) = &H1
    cboParity0.AddItem "Even", 2
    cboParity0.ItemData(2) = &H2
    
    cboParity1.AddItem "None", 0
    cboParity1.ItemData(0) = &H0
    cboParity1.AddItem "Odd", 1
    cboParity1.ItemData(1) = &H1
    cboParity1.AddItem "Even", 2
    cboParity1.ItemData(2) = &H2
    
    cboFlow0.AddItem "None", 0
    cboFlow0.ItemData(0) = &H0
    cboFlow0.AddItem "Xon/Xoff", 1
    cboFlow0.ItemData(1) = &H1
    cboFlow0.AddItem "CTS/RTS", 2
    cboFlow0.ItemData(2) = &H2
    
    cboFlow1.AddItem "None", 0
    cboFlow1.ItemData(0) = &H0
    cboFlow1.AddItem "Xon/Xoff", 1
    cboFlow1.ItemData(1) = &H1
    cboFlow1.AddItem "CTS/RTS", 2
    cboFlow1.ItemData(2) = &H2

    bSelect = False
    
    ToolMode = modeNone
    
    FrameTab(1).Left = TabStrip1.Left + 8
    FrameTab(1).Top = TabStrip1.Top + 24
    FrameTab(2).Left = TabStrip1.Left + 8
    FrameTab(2).Top = TabStrip1.Top + 24
    mintCurFrame = 1
    FrameTab(1).Visible = True
    
    txtDirectIP.Visible = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''
' Name : func_SearchBoard
' Parameter : None
'
' Search available Boards.
' Send "FIND" message
' Waiting Board's reply
'
''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub func_SearchBoard()

    Dim sendD() As Byte
    
    ' First, delete all board's information
    Call BoardRemove
    
    ToolMode = modeSearching
    ' Send FIND message
    ReDim sendD(0 To 3) As Byte
    sendD(0) = Asc("F")
    sendD(1) = Asc("I")
    sendD(2) = Asc("N")
    sendD(3) = Asc("D")
    
    If chkDirect.value = 1 Then
        WinsockDirect.RemoteHost = txtDirectIP.Text
        WinsockDirect.RemotePort = 1461
        WinsockDirect.Connect
        
    Else
        WinsockUDP.RemoteHost = "255.255.255.255"
        WinsockUDP.RemotePort = 1460
        WinsockUDP.SendData sendD
        Erase sendD
    End If
    
    Call ShowMsgWindow
    WinsockDirect.Close

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''
' Name : func_SettingBoard
' Parameter : None
'
' Update the selected Board's configuration data.
' Make message with new configuration data.
' Send "SETT" message
' Waiting Board's reply
'
''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub func_SettingBoard()

On Error GoTo s_ERROR

    Dim sendD() As Byte
    Dim tmpstr() As String
    Dim ii As Integer
    
  Dim Mystring As String
    
    
    ' If exist selected board ,
    If bSelect Then
        
        ' Getting selected board's information
        ' Making SETT message
        If chkDebug.value = 1 Then
            BoardInfo.debugoff = 0
        Else
            BoardInfo.debugoff = 1
        End If
        
        If Option1.value = True Then
            BoardInfo.DHCP = 0
        ElseIf Option2.value = True Then        ' DHCP
            BoardInfo.DHCP = 1
        ElseIf Option3.value = True Then    ' PPPoE
            BoardInfo.DHCP = 2
        End If
        
        BoardInfo.UDP0 = chkUDPMode0.value
        BoardInfo.UDP1 = chkUDPMode1.value
        BoardInfo.Connect0 = 0
        BoardInfo.Connect1 = 0
        
        tmpstr = Split(txtIP.Text, ".")
        If UBound(tmpstr) <> 3 Then
            Call MessageBox("Invalid IP Address.")
            txtIP.SetFocus
            Exit Sub
        End If
        For ii = 0 To 3
            If tmpstr(ii) = "" Or CInt(tmpstr(ii)) > 255 Or CInt(tmpstr(ii)) < 0 Then
                Call MessageBox("Invalid IP Address.")
                txtIP.SetFocus
                Exit Sub
            End If
            BoardInfo.ip(ii) = CByte(tmpstr(ii))
        Next ii
        
        tmpstr = Split(txtSubnet.Text, ".")
        If UBound(tmpstr) <> 3 Then
            Call MessageBox("Invalid Subnet Mask.")
            txtSubnet.SetFocus
            Exit Sub
        End If
        For ii = 0 To 3
            If tmpstr(ii) = "" Or CInt(tmpstr(ii)) > 255 Or CInt(tmpstr(ii)) < 0 Then
                Call MessageBox("Invalid Subnet Mask.")
                txtSubnet.SetFocus
                Exit Sub
            End If
            BoardInfo.subnet(ii) = CByte(tmpstr(ii))
        Next ii
    
        tmpstr = Split(txtGW.Text, ".")
        If UBound(tmpstr) <> 3 Then
            Call MessageBox("Invalid Gateway Address.")
            txtGW.SetFocus
            Exit Sub
        End If
        For ii = 0 To 3
            If tmpstr(ii) = "" Or CInt(tmpstr(ii)) > 255 Or CInt(tmpstr(ii)) < 0 Then
                Call MessageBox("Invalid Gateway Address.")
                txtGW.SetFocus
                Exit Sub
            End If
            BoardInfo.gw(ii) = CByte(tmpstr(ii))
        Next ii
        
        ' modified by JB
        BoardInfo.myport0(1) = (CLng(txtPort0.Text) And &HFF00) / &H100
        BoardInfo.myport0(0) = CLng(txtPort0.Text) And &HFF
        BoardInfo.myport1(1) = (CLng(txtPort1.Text) And &HFF00) / &H100
        BoardInfo.myport1(0) = CLng(txtPort1.Text) And &HFF
        
        If optClientMode0.Item(0).value Then
            BoardInfo.bserver0 = 0
        ElseIf optClientMode0.Item(1).value Then
            BoardInfo.bserver0 = 1
        Else
            BoardInfo.bserver0 = 2
        End If
        
        If optClientMode1.Item(0).value Then
            BoardInfo.bserver1 = 0
        ElseIf optClientMode1.Item(1).value Then
            BoardInfo.bserver1 = 1
        Else
            BoardInfo.bserver1 = 2
        End If
        
        tmpstr = Split(txtServerIP0.Text, ".")
        If UBound(tmpstr) <> 3 Then
            Call MessageBox("Invalid IP Address.")
            txtServerIP0.SetFocus
            Exit Sub
        End If
        For ii = 0 To 3
            If tmpstr(ii) = "" Or CInt(tmpstr(ii)) > 255 Or CInt(tmpstr(ii)) < 0 Then
                Call MessageBox("Invalid IP Address.")
                txtServerIP0.SetFocus
                Exit Sub
            End If
            BoardInfo.peerip0(ii) = CByte(tmpstr(ii))
        Next ii
        
        tmpstr = Split(txtServerIP1.Text, ".")
        If UBound(tmpstr) <> 3 Then
            Call MessageBox("Invalid IP Address.")
            txtServerIP1.SetFocus
            Exit Sub
        End If
        For ii = 0 To 3
            If tmpstr(ii) = "" Or CInt(tmpstr(ii)) > 255 Or CInt(tmpstr(ii)) < 0 Then
                Call MessageBox("Invalid IP Address.")
                txtServerIP1.SetFocus
                Exit Sub
            End If
            BoardInfo.peerip1(ii) = CByte(tmpstr(ii))
        Next ii
        
        '''''''''''' DNS
        ' UART 0
        BoardInfo.DNS_Flag0 = ChkDNS0.value
        
        If ChkDNS0.value = 1 Then
            tmpstr = Split(txt_DNS_ServerIP0.Text, ".")
            If UBound(tmpstr) <> 3 Then
                Call MessageBox("Invalid DNS Server IP Address.")
                txt_DNS_ServerIP0.SetFocus
                Exit Sub
            End If
            For ii = 0 To 3
                If tmpstr(ii) = "" Or CInt(tmpstr(ii)) > 255 Or CInt(tmpstr(ii)) < 0 Then
                    Call MessageBox("Invalid DNS Server IP Address.")
                    txt_DNS_ServerIP0.SetFocus
                    Exit Sub
                End If
                BoardInfo.DNS_IP0(ii) = CByte(tmpstr(ii))
            Next ii
            
            Mystring = txtServer_Domain0.Text
            ReDim mybytearray(0 To Len(Mystring) - 1) As Byte
            mybytearray() = StrConv(Mystring, vbFromUnicode)
            For ii = 0 To UBound(mybytearray)
              BoardInfo.D_SIP0(ii) = mybytearray(ii)
            Next ii
            For ii = UBound(mybytearray) + 1 To 31
              BoardInfo.D_SIP0(ii) = 0
            Next ii
        End If
        
        'UART 1
        BoardInfo.DNS_Flag1 = ChkDNS1.value
        
        If ChkDNS1.value = 1 Then
            tmpstr = Split(txt_DNS_ServerIP1.Text, ".")
            If UBound(tmpstr) <> 3 Then
                Call MessageBox("Invalid DNS Server IP Address.")
                txt_DNS_ServerIP1.SetFocus
                Exit Sub
            End If
            For ii = 0 To 3
                If tmpstr(ii) = "" Or CInt(tmpstr(ii)) > 255 Or CInt(tmpstr(ii)) < 0 Then
                    Call MessageBox("Invalid DNS Server IP Address.")
                    txt_DNS_ServerIP1.SetFocus
                    Exit Sub
                End If
                BoardInfo.DNS_IP1(ii) = CByte(tmpstr(ii))
            Next ii
            
            Mystring = txtServer_Domain1.Text
            ReDim mybytearray(0 To Len(Mystring) - 1) As Byte
            mybytearray() = StrConv(Mystring, vbFromUnicode)
            For ii = 0 To UBound(mybytearray)
              BoardInfo.D_SIP1(ii) = mybytearray(ii)
            Next ii
            For ii = UBound(mybytearray) + 1 To 31
              BoardInfo.D_SIP1(ii) = 0
            Next ii
        End If
    
        BoardInfo.peerport0(1) = (CLng(txtServerPort0.Text) And &HFF00) / &H100
        BoardInfo.peerport0(0) = CLng(txtServerPort0.Text) And &HFF
        BoardInfo.peerport1(1) = (CLng(txtServerPort1.Text) And &HFF00) / &H100
        BoardInfo.peerport1(0) = CLng(txtServerPort1.Text) And &HFF
        BoardInfo.I_time0(1) = (CLng(txtITime0.Text) And &HFF00) / &H100
        BoardInfo.I_time0(0) = CLng(txtITime0.Text) And &HFF
        BoardInfo.I_time1(1) = (CLng(txtITime1.Text) And &HFF00) / &H100
        BoardInfo.I_time1(0) = CLng(txtITime1.Text) And &HFF
        BoardInfo.D_time0(1) = (CLng(txtDTime0.Text) And &HFF00) / &H100
        BoardInfo.D_time0(0) = CLng(txtDTime0.Text) And &HFF
        BoardInfo.D_time1(1) = (CLng(txtDTime1.Text) And &HFF00) / &H100
        BoardInfo.D_time1(0) = CLng(txtDTime1.Text) And &HFF
        BoardInfo.D_size0(1) = (CInt(txtDSize0.Text) And &HFF00) / &H100
        BoardInfo.D_size0(0) = CInt(txtDSize0.Text) And &HFF
        BoardInfo.D_size1(1) = (CInt(txtDSize1.Text) And &HFF00) / &H100
        BoardInfo.D_size1(0) = CInt(txtDSize1.Text) And &HFF
        BoardInfo.D_ch0 = CInt("&h" & txtDChar0.Text)
        BoardInfo.D_ch1 = CInt("&h" & txtDChar1.Text)
        
        BoardInfo.speed0 = cboSpeed0.ItemData(cboSpeed0.ListIndex)
        BoardInfo.speed1 = cboSpeed0.ItemData(cboSpeed1.ListIndex)
        BoardInfo.databit0 = cboDataBits0.ItemData(cboDataBits0.ListIndex)
        BoardInfo.databit1 = cboDataBits1.ItemData(cboDataBits1.ListIndex)
        BoardInfo.parity0 = cboParity0.ItemData(cboParity0.ListIndex)
        BoardInfo.parity1 = cboParity1.ItemData(cboParity1.ListIndex)
        BoardInfo.stopbit0 = cboStopBits0.ItemData(cboStopBits0.ListIndex)
        BoardInfo.stopbit1 = cboStopBits1.ItemData(cboStopBits1.ListIndex)
        BoardInfo.flow0 = cboFlow0.ItemData(cboFlow0.ListIndex)
        BoardInfo.flow1 = cboFlow1.ItemData(cboFlow1.ListIndex)
        
        If CheckSCfg0.value = 1 Then
            BoardInfo.SCfg0 = 1
        Else
            BoardInfo.SCfg0 = 0
        End If
        
              
        BoardInfo.SCfgStr0(0) = CInt("&h" & TextSCfg0(0).Text)
        BoardInfo.SCfgStr0(1) = CInt("&h" & TextSCfg0(1).Text)
        BoardInfo.SCfgStr0(2) = CInt("&h" & TextSCfg0(2).Text)
        
        If Option3.value = True Then
            Mystring = txtID.Text
            ReDim mybytearray(0 To Len(Mystring) - 1) As Byte
            mybytearray() = StrConv(Mystring, vbFromUnicode)
            For ii = 0 To UBound(mybytearray)
              BoardInfo.PPPoE_ID(ii) = mybytearray(ii)
            Next ii
            For ii = UBound(mybytearray) + 1 To 31
              BoardInfo.PPPoE_ID(ii) = 0
            Next ii
            
            Mystring = txtPass.Text
            ReDim mybytearray(0 To Len(Mystring) - 1) As Byte
            mybytearray() = StrConv(Mystring, vbFromUnicode)
            For ii = 0 To UBound(mybytearray)
              BoardInfo.PPPoE_Pass(ii) = mybytearray(ii)
            Next ii
            For ii = UBound(mybytearray) + 1 To 31
              BoardInfo.PPPoE_Pass(ii) = 0
            Next ii
        End If

        If chkTCPPass0.value = 1 Then
            BoardInfo.EnTCPPass0 = 1
            
            Mystring = txtTCPPass0.Text
            If Len(Mystring) > 8 Then
                ReDim mybytearray(0 To 7) As Byte
            Else
                ReDim mybytearray(0 To Len(Mystring) - 1) As Byte
            End If
            
            mybytearray() = StrConv(Mystring, vbFromUnicode)
            For ii = 0 To UBound(mybytearray)
              BoardInfo.TCPPass0(ii) = mybytearray(ii)
            Next ii
            For ii = UBound(mybytearray) + 1 To 7
              BoardInfo.TCPPass0(ii) = 0
            Next ii
        Else
            BoardInfo.EnTCPPass0 = 0
            For ii = 0 To 7
                BoardInfo.TCPPass0(ii) = 0
            Next ii
        End If
        
        If chkTCPPass1.value = 1 Then
            BoardInfo.EnTCPPass1 = 1
            
            Mystring = txtTCPPass1.Text
            If Len(Mystring) > 8 Then
                ReDim mybytearray(0 To 7) As Byte
            Else
                ReDim mybytearray(0 To Len(Mystring) - 1) As Byte
            End If
            
            mybytearray() = StrConv(Mystring, vbFromUnicode)
            For ii = 0 To UBound(mybytearray)
              BoardInfo.TCPPass1(ii) = mybytearray(ii)
            Next ii
            For ii = UBound(mybytearray) + 1 To 7
              BoardInfo.TCPPass1(ii) = 0
            Next ii
        Else
            BoardInfo.EnTCPPass1 = 0
            For ii = 0 To 7
                BoardInfo.TCPPass1(ii) = 0
            Next ii
        End If
        
        
        ToolMode = modeSetting

        ' Sending SETT message
        ReDim sendD(0 To Len(BoardInfo) + 3) As Byte
        sendD(0) = Asc("S")
        sendD(1) = Asc("E")
        sendD(2) = Asc("T")
        sendD(3) = Asc("T")
        CopyMemory sendD(4), BoardInfo, Len(BoardInfo)
        
        If chkDirect.value = 1 Then
            WinsockDirect.RemoteHost = txtDirectIP.Text
            WinsockDirect.RemotePort = 1461
            WinsockDirect.Connect
            
        Else
            WinsockUDP.RemoteHost = "255.255.255.255"
            WinsockUDP.RemotePort = 1460
            WinsockUDP.SendData sendD
            Erase sendD
        End If
        
        Call ShowMsgWindow
        WinsockDirect.Close
    
    End If
    Exit Sub
    
    '추가
    'FrmTcpServer.txt_LPRIP(
    
    
s_ERROR:
    Call MessageBox("Invalid parameter value.")
    
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''
' Name : func_Upload
' Parameter : None
'
' Uploading new firmware to selected Board.
' Send "FIRS" message for alert uploading to selected Board.
' Try to connect for making uploading socket.
' waiting until connecting Board.
'
''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub func_Upload()

    On Error Resume Next
    
    Dim Ret As Integer      ' Return value
    Dim sendD() As Byte
    Dim tmpstr() As String
    Dim ii As Integer
    
    If bSelect Then
        
        tmpstr = Split(txtIP.Text, ".")
        If UBound(tmpstr) <> 3 Then
            Call MessageBox("Invalid IP Address.")
            txtIP.SetFocus
            Exit Sub
        End If
        For ii = 0 To 3
            If tmpstr(ii) = "" Or CInt(tmpstr(ii)) > 255 Or CInt(tmpstr(ii)) < 0 Then
                Call MessageBox("Invalid IP Address.")
                txtIP.SetFocus
                Exit Sub
            End If
        Next ii
        
        ' Select firmware's file
        OpenLog.DialogTitle = "File Select"
        OpenLog.Filter = "Bin File (*.bin)|*.bin|All File (*.*)|*.*"
        Do
            OpenLog.CancelError = True
            OpenLog.filename = ""
            OpenLog.ShowOpen
            If Err = cdlCancel Then
                Exit Sub
            End If
            
            strUploadFile = OpenLog.filename
            ' if file not exist, return.
            Ret = Len(Dir$(strUploadFile))
            If Err Then
               Call MessageBox(Error$)
               Exit Sub
            End If
            If Ret Then
               Exit Do
            Else
               Call MessageBox("No existing " + strUploadFile)
            End If
        Loop
        
        ToolMode = modeUploading
        bDirectUpload = False
        
        ' Inform board uploading
        ' Send FIRS message
        ReDim sendD(0 To Len(BoardInfo) + 3) As Byte
        sendD(0) = Asc("F")
        sendD(1) = Asc("I")
        sendD(2) = Asc("R")
        sendD(3) = Asc("S")
        CopyMemory sendD(4), BoardInfo, Len(BoardInfo)
        
        If chkDirect.value = 1 Then
            WinsockDirect.RemoteHost = txtDirectIP.Text
            WinsockDirect.RemotePort = 1461
            WinsockDirect.Connect
            
            While Not bDirectUpload
               DoEvents
            Wend
            
        Else
            WinsockUDP.RemoteHost = "255.255.255.255"
            WinsockUDP.RemotePort = 1460
            WinsockUDP.SendData sendD
            Erase sendD
        End If
        
        
        If chkDirect.value = 1 Then
            destIP = txtDirectIP.Text
            Sleep (100)
        Else
            destIP = txtIP.Text
            Sleep (1)
        End If
        
        'destIP = txtIP.Text
        'destIP = "211.171.137.58"
        Call ShowMsgWindow
        WinsockDirect.Close
        
    End If

End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.KEY
        Case "SearchBoard"
            Call func_SearchBoard
        Case "SettingBoard"
            Call func_SettingBoard
        Case "Upload"
            Call func_Upload
        Case "Exit"
            'Form_Unload 0
            Me.Hide
    End Select
End Sub

Public Sub ListBoards_FirstRowSelect()
    ListBoards_ItemClick ListBoards.ListItems(1)
    ListBoards.ListItems(1).Selected = True
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''
' Name : ListBoards_ItemClick
' Parameter : Item is string of selected board's key.
'
' Save key string to "BoardKey" variable
' Save configuration data to "BoardInfo" variable.
' Show configuration data of selected Board.
'
''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub ListBoards_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Dim DD() As Byte
    Dim i As Long
    
On Error GoTo click_ERROR
    
    bSelect = True
    DD = colBoards.Item(Item.KEY)
    BoardKey = Item.KEY
    CopyMemory BoardInfo, DD(0), Len(BoardInfo)
    
    txtVersion.Text = BoardInfo.AppVer(0) & "." & BoardInfo.AppVer(1)
    If BoardInfo.debugoff = 0 Then
        chkDebug.value = 1
    Else
        chkDebug.value = 0
    End If
    
    If BoardInfo.DHCP = 1 Then
        Option2.value = 1
    ElseIf BoardInfo.DHCP = 2 Then
        Option3.value = 1
    ElseIf BoardInfo.DHCP = 0 Then
        Option1.value = 1
    End If
    
    txtID.Text = StrConv(BoardInfo.PPPoE_ID, vbUnicode)
    txtPass.Text = StrConv(BoardInfo.PPPoE_Pass, vbUnicode)

    If BoardInfo.UDP0 = 1 Then
        chkUDPMode0.value = 1
    Else
        chkUDPMode0.value = 0
    End If
    
    If BoardInfo.UDP1 = 1 Then
        chkUDPMode1.value = 1
    Else
        chkUDPMode1.value = 0
    End If
    
    If BoardInfo.Connect0 = 1 Then
        txtConnect0.Text = "Connected"
    Else
        txtConnect0.Text = "Not Connected"
    End If
    
    If BoardInfo.Connect1 = 1 Then
        txtConnect1.Text = "Connected"
    Else
        txtConnect1.Text = "Not Connected"
    End If
        
    txtIP.Text = BoardInfo.ip(0) & "." & BoardInfo.ip(1) & "." & BoardInfo.ip(2) & "." & BoardInfo.ip(3)
    txtSubnet.Text = BoardInfo.subnet(0) & "." & BoardInfo.subnet(1) & "." & BoardInfo.subnet(2) & "." & BoardInfo.subnet(3)
    txtGW.Text = BoardInfo.gw(0) & "." & BoardInfo.gw(1) & "." & BoardInfo.gw(2) & "." & BoardInfo.gw(3)
    i = BoardInfo.myport0(1)
    i = (i * &H100)
    i = i + BoardInfo.myport0(0)
    txtPort0.Text = CStr(i)
    
    i = BoardInfo.myport1(1)
    i = (i * &H100)
    i = i + BoardInfo.myport1(0)
    txtPort1.Text = CStr(i)
    
    'Added Ver1.1. for the operation mode
            
    If BoardInfo.bserver0 = 0 Then
        optClientMode0.Item(0).value = True
        optClientMode0.Item(1).value = False
        optClientMode0.Item(2).value = False
    ElseIf BoardInfo.bserver0 = 1 Then
        optClientMode0.Item(0).value = False
        optClientMode0.Item(1).value = True
        optClientMode0.Item(2).value = False
    Else
        optClientMode0.Item(0).value = False
        optClientMode0.Item(1).value = False
        optClientMode0.Item(2).value = True
    End If
    If BoardInfo.bserver1 = 0 Then
        optClientMode1.Item(0).value = True
        optClientMode1.Item(1).value = False
        optClientMode1.Item(2).value = False
    ElseIf BoardInfo.bserver1 = 1 Then
        optClientMode1.Item(0).value = False
        optClientMode1.Item(1).value = True
        optClientMode1.Item(2).value = False
    Else
        optClientMode1.Item(0).value = False
        optClientMode1.Item(1).value = False
        optClientMode1.Item(2).value = True
    End If
    ' End of addition Ver1.1
    
        
        
        
    ''DNS
    txt_DNS_ServerIP0.Text = BoardInfo.DNS_IP0(0) & "." & BoardInfo.DNS_IP0(1) & "." & BoardInfo.DNS_IP0(2) & "." & BoardInfo.DNS_IP0(3)
    txtServer_Domain0.Text = StrConv(BoardInfo.D_SIP0, vbUnicode)
    If BoardInfo.DNS_Flag0 = 1 Then
       ChkDNS0.value = 1
    Else
       ChkDNS0.value = 0
    End If
    
    txt_DNS_ServerIP1.Text = BoardInfo.DNS_IP1(0) & "." & BoardInfo.DNS_IP1(1) & "." & BoardInfo.DNS_IP1(2) & "." & BoardInfo.DNS_IP1(3)
    txtServer_Domain1.Text = StrConv(BoardInfo.D_SIP1, vbUnicode)
    If BoardInfo.DNS_Flag1 = 1 Then
       ChkDNS1.value = 1
    Else
       ChkDNS1.value = 0
    End If
    
    'Destination Ip
    txtServerIP0.Text = BoardInfo.peerip0(0) & "." & BoardInfo.peerip0(1) & "." & BoardInfo.peerip0(2) & "." & BoardInfo.peerip0(3)
    i = BoardInfo.peerport0(1)
    i = (i * &H100)
    i = i + BoardInfo.peerport0(0)
    txtServerPort0.Text = CStr(i)
    
    txtServerIP1.Text = BoardInfo.peerip1(0) & "." & BoardInfo.peerip1(1) & "." & BoardInfo.peerip1(2) & "." & BoardInfo.peerip1(3)
    i = BoardInfo.peerport1(1)
    i = (i * &H100)
    i = i + BoardInfo.peerport1(0)
    txtServerPort1.Text = CStr(i)
    
    i = BoardInfo.I_time0(1)
    i = (i * &H100)
    i = i + BoardInfo.I_time0(0)
    txtITime0.Text = CStr(i)
    
    i = BoardInfo.I_time1(1)
    i = (i * &H100)
    i = i + BoardInfo.I_time1(0)
    txtITime1.Text = CStr(i)
    
    i = BoardInfo.D_time0(1)
    i = (i * &H100)
    i = i + BoardInfo.D_time0(0)
    txtDTime0.Text = CStr(i)
    
    i = BoardInfo.D_time1(1)
    i = (i * &H100)
    i = i + BoardInfo.D_time1(0)
    txtDTime1.Text = CStr(i)
    
    i = (BoardInfo.D_size0(1) * &H100) + BoardInfo.D_size0(0)
    txtDSize0.Text = CStr(i)
    i = (BoardInfo.D_size1(1) * &H100) + BoardInfo.D_size1(0)
    txtDSize1.Text = CStr(i)
    
    If BoardInfo.D_ch0 > 15 Then
        txtDChar0.Text = Hex(BoardInfo.D_ch0)
    Else
        txtDChar0.Text = "0" & Hex(BoardInfo.D_ch0)
    End If
    
    If BoardInfo.D_ch1 > 15 Then
        txtDChar1.Text = Hex(BoardInfo.D_ch1)
    Else
        txtDChar1.Text = "0" & Hex(BoardInfo.D_ch1)
    End If
    
    
    Select Case BoardInfo.speed0
        Case &HBB
            cboSpeed0.ListIndex = 8
        Case &HFF
            cboSpeed0.ListIndex = 7
        Case &HFE
            cboSpeed0.ListIndex = 6
        Case &HFD
            cboSpeed0.ListIndex = 5
        Case &HFA
            cboSpeed0.ListIndex = 4
        Case &HF4
            cboSpeed0.ListIndex = 3
        Case &HE8
            cboSpeed0.ListIndex = 2
        Case &HD0
            cboSpeed0.ListIndex = 1
        Case &HA0
            cboSpeed0.ListIndex = 0
        Case Else
            cboSpeed0.ListIndex = 0
    End Select
    
    Select Case BoardInfo.speed1
        Case &HBB
            cboSpeed1.ListIndex = 8
        Case &HFF
            cboSpeed1.ListIndex = 7
        Case &HFE
            cboSpeed1.ListIndex = 6
        Case &HFD
            cboSpeed1.ListIndex = 5
        Case &HFA
            cboSpeed1.ListIndex = 4
        Case &HF4
            cboSpeed1.ListIndex = 3
        Case &HE8
            cboSpeed1.ListIndex = 2
        Case &HD0
            cboSpeed1.ListIndex = 1
        Case &HA0
            cboSpeed1.ListIndex = 0
        Case Else
            cboSpeed1.ListIndex = 0
    End Select
    
    cboDataBits0.Text = CStr(BoardInfo.databit0)
    cboDataBits1.Text = CStr(BoardInfo.databit1)
    cboStopBits0.Text = CStr(BoardInfo.stopbit0)
    cboStopBits1.Text = CStr(BoardInfo.stopbit1)
    
    Select Case BoardInfo.parity0
        Case &H0
            cboParity0.Text = "None"
        Case &H1
            cboParity0.Text = "Odd"
        Case &H2
            cboParity0.Text = "Even"
        Case Else
            cboParity0.Text = "None"
    End Select
    Select Case BoardInfo.parity1
        Case &H0
            cboParity1.Text = "None"
        Case &H1
            cboParity1.Text = "Odd"
        Case &H2
            cboParity1.Text = "Even"
        Case Else
            cboParity1.Text = "None"
    End Select
    cboFlow0.ListIndex = BoardInfo.flow0
    cboFlow1.ListIndex = BoardInfo.flow1
    
   ' Serial Configure
    CheckSCfg0.Enabled = True
    TextSCfg0(0).Enabled = True
    TextSCfg0(1).Enabled = True
    TextSCfg0(2).Enabled = True
    
    If BoardInfo.SCfg0 = 0 Then
        CheckSCfg0.value = 0
    Else
        CheckSCfg0.value = 1
    End If
        
        
    If BoardInfo.SCfgStr0(0) > 15 Then
        TextSCfg0(0).Text = Hex(BoardInfo.SCfgStr0(0))
    Else
        TextSCfg0(0).Text = "0" & Hex(BoardInfo.SCfgStr0(0))
    End If
    If BoardInfo.SCfgStr0(1) > 15 Then
        TextSCfg0(1).Text = Hex(BoardInfo.SCfgStr0(1))
    Else
        TextSCfg0(1).Text = "0" & Hex(BoardInfo.SCfgStr0(1))
    End If
    If BoardInfo.SCfgStr0(2) > 15 Then
        TextSCfg0(2).Text = Hex(BoardInfo.SCfgStr0(2))
    Else
        TextSCfg0(2).Text = "0" & Hex(BoardInfo.SCfgStr0(2))
    End If
        
    'TCP password configuration
    'Added ver1.1
    chkTCPPass0.Enabled = True
    txtTCPPass0.Enabled = True
    
    If BoardInfo.EnTCPPass0 = 0 Then
        chkTCPPass0.value = 0
    Else
        chkTCPPass0.value = 1
    End If
    
    txtTCPPass0.Text = StrConv(BoardInfo.TCPPass0, vbUnicode)
    
    chkTCPPass1.Enabled = True
    txtTCPPass1.Enabled = True
    
    If BoardInfo.EnTCPPass1 = 0 Then
        chkTCPPass1.value = 0
    Else
        chkTCPPass1.value = 1
    End If
    
    txtTCPPass1.Text = StrConv(BoardInfo.TCPPass1, vbUnicode)
    ' End of addition ver1.1
            
    Erase DD
    Exit Sub

click_ERROR:
    cboFlow0.ListIndex = 0
    cboFlow1.ListIndex = 0
    Call MessageBox("Invalid parameter's value.")
    
End Sub

Private Sub cboDataBits0_Click()
    If cboDataBits0.Text = "7" Then
        cboParity0.Clear
        'cboParity0.AddItem "None", 0
        'cboParity0.ItemData(0) = &H0
        cboParity0.AddItem "Odd", 0
        cboParity0.ItemData(0) = &H1
        cboParity0.AddItem "Even", 1
        cboParity0.ItemData(1) = &H2
        cboParity0.ListIndex = 0
    Else
        cboParity0.Clear
        cboParity0.AddItem "None", 0
        cboParity0.ItemData(0) = &H0
        cboParity0.AddItem "Odd", 1
        cboParity0.ItemData(1) = &H1
        cboParity0.AddItem "Even", 2
        cboParity0.ItemData(2) = &H2
        cboParity0.ListIndex = 0
    End If

End Sub
Private Sub cboDataBits1_Click()
    If cboDataBits1.Text = "7" Then
        cboParity1.Clear
        'cboParity.AddItem "None", 0
        'cboParity.ItemData(0) = &H0
        cboParity1.AddItem "Odd", 0
        cboParity1.ItemData(0) = &H1
        cboParity1.AddItem "Even", 1
        cboParity1.ItemData(1) = &H2
        cboParity1.ListIndex = 0
    Else
        cboParity1.Clear
        cboParity1.AddItem "None", 0
        cboParity1.ItemData(0) = &H0
        cboParity1.AddItem "Odd", 1
        cboParity1.ItemData(1) = &H1
        cboParity1.AddItem "Even", 2
        cboParity1.ItemData(2) = &H2
        cboParity1.ListIndex = 0
    End If

End Sub



Private Sub WinsockDirect_Connect()
    Dim sendD() As Byte
    
    
    Select Case ToolMode
    Case modeSearching
            ReDim sendD(0 To 3) As Byte
            sendD(0) = Asc("F")
            sendD(1) = Asc("I")
            sendD(2) = Asc("N")
            sendD(3) = Asc("D")
    
    Case modeSetting
            ' Sending SETT message
            ReDim sendD(0 To Len(BoardInfo) + 3) As Byte
            sendD(0) = Asc("S")
            sendD(1) = Asc("E")
            sendD(2) = Asc("T")
            sendD(3) = Asc("T")
            CopyMemory sendD(4), BoardInfo, Len(BoardInfo)
    
    Case modeUploading
            ReDim sendD(0 To Len(BoardInfo) + 3) As Byte
            sendD(0) = Asc("F")
            sendD(1) = Asc("I")
            sendD(2) = Asc("R")
            sendD(3) = Asc("S")
            CopyMemory sendD(4), BoardInfo, Len(BoardInfo)
    Case Else
        WinsockDirect.Close
        Exit Sub
    End Select
    
    WinsockDirect.SendData sendD
    Erase sendD
    
    bDirectUpload = True

End Sub

Private Sub WinsockDirect_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    WinsockDirect.Close

End Sub

Private Sub WinsockDirect_DataArrival(ByVal bytesTotal As Long)
On Error GoTo WinsockDirect_DataArrival_ERROR
'On Error Resume Next
    Dim getd() As Byte
    Dim getboard(BoardInfoSize_3_4) As Byte
    Dim getkind(3) As Byte
    
    ReDim getd(0 To (bytesTotal - 1)) As Byte
    WinsockDirect.GetData getd, vbByte, bytesTotal
    CopyMemory getkind(0), getd(0), 4
    CopyMemory getboard(0), getd(4), BoardInfoSize_3_4 - 4
    Erase getd
    
    If (getkind(0) = Asc("I")) And (getkind(1) = Asc("M")) And (getkind(2) = Asc("I")) And (getkind(3) = Asc("N")) Then
        If ToolMode = modeSearching Then
            Call BoardAdd(getboard)
            'ToolMode = None
        End If
        
    ElseIf (getkind(0) = Asc("S")) And (getkind(1) = Asc("E")) And (getkind(2) = Asc("T")) And (getkind(3) = Asc("C")) Then
        If ToolMode = modeSetting Then
            Call BoardUpdate(getboard)
        End If
    End If
    
    WinsockDirect.Close
    
    Exit Sub

WinsockDirect_DataArrival_ERROR:
    If Err Then
'        MsgBox "Retry, please"
        MsgBox Err.Description
    End If
    Erase getd

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''
' Name : WinsockUDP_DataArrival
' Parameter : bytesTotal is count of receiving data from "WinsockUDP" control
'
' Receive configuration message.
' "IMIN" message => BoardAdd
' "SETC" message => BoardUpdate
'
''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub WinsockUDP_DataArrival(ByVal bytesTotal As Long)
On Error GoTo WinsockUDP_DataArrival_ERROR
'On Error Resume Next
    Dim getd() As Byte
    Dim getboard(BoardInfoSize_3_4) As Byte
    Dim getkind(3) As Byte
    
    ReDim getd(0 To (bytesTotal - 1)) As Byte
    WinsockUDP.GetData getd, vbByte, bytesTotal
    CopyMemory getkind(0), getd(0), 4
    CopyMemory getboard(0), getd(4), BoardInfoSize_3_4 - 4
    Erase getd
    
    If (getkind(0) = Asc("I")) And (getkind(1) = Asc("M")) And (getkind(2) = Asc("I")) And (getkind(3) = Asc("N")) Then
        If ToolMode = modeSearching Then
            Call BoardAdd(getboard)
            'ToolMode = None
        End If
        
    ElseIf (getkind(0) = Asc("S")) And (getkind(1) = Asc("E")) And (getkind(2) = Asc("T")) And (getkind(3) = Asc("C")) Then
        If ToolMode = modeSetting Then
            Call BoardUpdate(getboard)
        End If
    End If
    Exit Sub

WinsockUDP_DataArrival_ERROR:
    If Err Then
'        MsgBox "Retry, please"
        MsgBox Err.Description
    End If
    Erase getd
    
End Sub

Private Sub chkDirect_Click()
    If chkDirect.value = 1 Then
        txtDirectIP.Visible = True
    Else
        txtDirectIP.Visible = False
    End If
    
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''
' Name : TEXT Box filtering
' Parameter : None
'
' Filtering textbox's data
' ex) Port text box support only number
'
''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub txtVersion_KeyPress(KeyAscii As Integer)
    
        KeyAscii = 0

End Sub

Private Sub txtDChar0_KeyPress(KeyAscii As Integer)
    
    If (KeyAscii = 8) Or (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 70) Or (KeyAscii >= 97 And KeyAscii <= 102) Then
    ' backspace or 0~9 or A~F or a~f
    Else
    ' else ignore
        KeyAscii = 0
    End If

End Sub
Private Sub txtDChar1_KeyPress(KeyAscii As Integer)
    
    If (KeyAscii = 8) Or (KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii >= 65 And KeyAscii <= 70) Or (KeyAscii >= 97 And KeyAscii <= 102) Then
    ' backspace or 0~9 or A~F or a~f
    Else
    ' else ignore
        KeyAscii = 0
    End If

End Sub

Private Sub txtDTime0_KeyPress(KeyAscii As Integer)
    
    If (KeyAscii = 8) Or (KeyAscii >= 48 And KeyAscii <= 57) Then
    ' backspace or 0~9
    Else
    ' else ignore
        KeyAscii = 0
    End If

End Sub
Private Sub txtDTime1_KeyPress(KeyAscii As Integer)
    
    If (KeyAscii = 8) Or (KeyAscii >= 48 And KeyAscii <= 57) Then
    ' backspace or 0~9
    Else
    ' else ignore
        KeyAscii = 0
    End If

End Sub

Private Sub txtITime0_KeyPress(KeyAscii As Integer)
    
    If (KeyAscii = 8) Or (KeyAscii >= 48 And KeyAscii <= 57) Then
    ' backspace or 0~9
    Else
    ' else ignore
        KeyAscii = 0
    End If

End Sub
Private Sub txtITime1_KeyPress(KeyAscii As Integer)
    
    If (KeyAscii = 8) Or (KeyAscii >= 48 And KeyAscii <= 57) Then
    ' backspace or 0~9
    Else
    ' else ignore
        KeyAscii = 0
    End If

End Sub

Private Sub txtDSize0_KeyPress(KeyAscii As Integer)
    
    If (KeyAscii = 8) Or (KeyAscii >= 48 And KeyAscii <= 57) Then
    ' backspace or 0~9
    Else
    ' else ignore
        KeyAscii = 0
    End If

End Sub
Private Sub txtDSize1_KeyPress(KeyAscii As Integer)
    
    If (KeyAscii = 8) Or (KeyAscii >= 48 And KeyAscii <= 57) Then
    ' backspace or 0~9
    Else
    ' else ignore
        KeyAscii = 0
    End If

End Sub

Private Sub txtIP_KeyPress(KeyAscii As Integer)
    
    If (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii >= 48 And KeyAscii <= 57) Then
    ' backspace or . or 0~9
    Else
    ' else ignore
        KeyAscii = 0
    End If

End Sub

Private Sub txtDirectIP_KeyPress(KeyAscii As Integer)
    
    If (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii >= 48 And KeyAscii <= 57) Then
    ' backspace or . or 0~9
    Else
    ' else ignore
        KeyAscii = 0
    End If

End Sub

Private Sub txtServerIP0_KeyPress(KeyAscii As Integer)
    
    If (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii >= 48 And KeyAscii <= 57) Then
    ' backspace or . or 0~9
    Else
    ' else ignore
        KeyAscii = 0
    End If

End Sub
Private Sub txtServerIP1_KeyPress(KeyAscii As Integer)
    
    If (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii >= 48 And KeyAscii <= 57) Then
    ' backspace or . or 0~9
    Else
    ' else ignore
        KeyAscii = 0
    End If

End Sub

Private Sub txtPort0_KeyPress(KeyAscii As Integer)
    
    If (KeyAscii = 8) Or (KeyAscii >= 48 And KeyAscii <= 57) Then
    ' backspace or 0~9
    Else
    ' else ignore
        KeyAscii = 0
    End If

End Sub
Private Sub txtPort1_KeyPress(KeyAscii As Integer)
    
    If (KeyAscii = 8) Or (KeyAscii >= 48 And KeyAscii <= 57) Then
    ' backspace or 0~9
    Else
    ' else ignore
        KeyAscii = 0
    End If

End Sub

Private Sub txtServerPort0_KeyPress(KeyAscii As Integer)
    
    If (KeyAscii = 8) Or (KeyAscii >= 48 And KeyAscii <= 57) Then
    ' backspace or 0~9
    Else
    ' else ignore
        KeyAscii = 0
    End If

End Sub
Private Sub txtServerPort1_KeyPress(KeyAscii As Integer)
    
    If (KeyAscii = 8) Or (KeyAscii >= 48 And KeyAscii <= 57) Then
    ' backspace or 0~9
    Else
    ' else ignore
        KeyAscii = 0
    End If

End Sub

Private Sub txtConnect0_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub txtConnect1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txt_DNS_ServerIP0_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii >= 48 And KeyAscii <= 57) Then
    ' backspace or . or 0~9
    Else
    ' else ignore
        KeyAscii = 0
    End If
End Sub
Private Sub txt_DNS_ServerIP1_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 8) Or (KeyAscii = 46) Or (KeyAscii >= 48 And KeyAscii <= 57) Then
    ' backspace or . or 0~9
    Else
    ' else ignore
        KeyAscii = 0
    End If
End Sub

