VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmCouponSale 
   Caption         =   " ���α� �Ǹ� / ����"
   ClientHeight    =   14655
   ClientLeft      =   4080
   ClientTop       =   1440
   ClientWidth     =   19080
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "FrmCouponSale.frx":0000
   ScaleHeight     =   14655
   ScaleWidth      =   19080
   Begin VB.ComboBox cmb_Sort 
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
      ItemData        =   "FrmCouponSale.frx":F4A3
      Left            =   14850
      List            =   "FrmCouponSale.frx":F4B0
      TabIndex        =   43
      Text            =   "cmb_Sort"
      Top             =   2160
      Width           =   1545
   End
   Begin VB.ComboBox cmb_Gubun 
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      ItemData        =   "FrmCouponSale.frx":F4CC
      Left            =   6930
      List            =   "FrmCouponSale.frx":F4CE
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   2
      Top             =   10140
      Width           =   2325
   End
   Begin VB.TextBox txt_Dong 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   11520
      TabIndex        =   4
      Top             =   10110
      Width           =   1155
   End
   Begin VB.TextBox txt_Ho 
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13380
      TabIndex        =   5
      Top             =   10110
      Width           =   1155
   End
   Begin VB.TextBox txt_Phone 
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2700
      TabIndex        =   1
      Top             =   11070
      Width           =   2325
   End
   Begin VB.TextBox txt_Name 
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2700
      TabIndex        =   0
      Top             =   10605
      Width           =   2325
   End
   Begin VB.TextBox txt_Num 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2700
      TabIndex        =   29
      Top             =   10140
      Width           =   2325
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "�������"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1200
      Left            =   120
      TabIndex        =   12
      Top             =   13350
      Width           =   18975
   End
   Begin VB.TextBox txt_tmpName 
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
      Left            =   4590
      TabIndex        =   11
      Text            =   "SDFS"
      Top             =   2160
      Width           =   2115
   End
   Begin ComctlLib.ListView ListView_REG 
      Height          =   5115
      Left            =   360
      TabIndex        =   13
      Top             =   3720
      Width           =   18450
      _ExtentX        =   32544
      _ExtentY        =   9022
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
      Height          =   660
      Index           =   0
      Left            =   17580
      TabIndex        =   14
      Top             =   690
      Width           =   1350
      _Version        =   65536
      _ExtentX        =   2381
      _ExtentY        =   1164
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
      Picture         =   "FrmCouponSale.frx":F4D0
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   660
      Index           =   1
      Left            =   16140
      TabIndex        =   10
      Top             =   690
      Width           =   1350
      _Version        =   65536
      _ExtentX        =   2381
      _ExtentY        =   1164
      _StockProps     =   78
      Caption         =   "Excel"
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
      Picture         =   "FrmCouponSale.frx":F821
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   660
      Index           =   2
      Left            =   16920
      TabIndex        =   15
      Top             =   2040
      Width           =   1650
      _Version        =   65536
      _ExtentX        =   2910
      _ExtentY        =   1164
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
      Picture         =   "FrmCouponSale.frx":FB72
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   405
      Left            =   8310
      TabIndex        =   16
      Top             =   2160
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�������"
         Size            =   11.25
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
      Format          =   53739520
      CurrentDate     =   36927
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   405
      Left            =   11490
      TabIndex        =   17
      Top             =   2160
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   714
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�������"
         Size            =   11.25
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
      Format          =   53739520
      CurrentDate     =   36927
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   570
      Index           =   3
      Left            =   17010
      TabIndex        =   8
      Top             =   11340
      Width           =   1350
      _Version        =   65536
      _ExtentX        =   2381
      _ExtentY        =   1005
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
      Picture         =   "FrmCouponSale.frx":FEC3
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   570
      Index           =   4
      Left            =   15600
      TabIndex        =   7
      Top             =   11340
      Width           =   1350
      _Version        =   65536
      _ExtentX        =   2381
      _ExtentY        =   1005
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
      Picture         =   "FrmCouponSale.frx":10214
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   570
      Index           =   5
      Left            =   14175
      TabIndex        =   6
      Top             =   11340
      Width           =   1350
      _Version        =   65536
      _ExtentX        =   2381
      _ExtentY        =   1005
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
      Picture         =   "FrmCouponSale.frx":10565
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   570
      Index           =   6
      Left            =   12750
      TabIndex        =   9
      Top             =   11340
      Width           =   1350
      _Version        =   65536
      _ExtentX        =   2381
      _ExtentY        =   1005
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
      Picture         =   "FrmCouponSale.frx":108B6
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
      Index           =   0
      Left            =   6930
      TabIndex        =   30
      Top             =   10620
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
      Index           =   1
      Left            =   6930
      TabIndex        =   40
      Top             =   11580
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
      Index           =   2
      Left            =   6930
      TabIndex        =   3
      Top             =   11100
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
      Alignment       =   1  '������ ����
      BackStyle       =   0  '����
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   15.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   17580
      TabIndex        =   42
      Top             =   3240
      Width           =   405
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '����
      Caption         =   "�����ݾ�"
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   465
      Index           =   1
      Left            =   5580
      TabIndex        =   41
      Top             =   11595
      Width           =   1305
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '����
      Caption         =   "�� �� ��"
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   465
      Left            =   5580
      TabIndex        =   39
      Top             =   10140
      Width           =   1185
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '����
      Caption         =   "��     ��"
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   465
      Index           =   0
      Left            =   5580
      TabIndex        =   38
      Top             =   10635
      Width           =   1305
   End
   Begin VB.Label lbl_dept 
      BackStyle       =   0  '����
      Caption         =   "�����ȣ"
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   465
      Index           =   2
      Left            =   10170
      TabIndex        =   37
      Top             =   10140
      Width           =   1305
   End
   Begin VB.Label lbl_Phone 
      BackStyle       =   0  '����
      Caption         =   "��ȭ��ȣ"
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   465
      Left            =   1350
      TabIndex        =   36
      Top             =   11070
      Width           =   1305
   End
   Begin VB.Label lbl_dept 
      BackStyle       =   0  '����
      Caption         =   "~"
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   465
      Index           =   3
      Left            =   12960
      TabIndex        =   35
      Top             =   10125
      Width           =   225
   End
   Begin VB.Label lbl_Num 
      BackStyle       =   0  '����
      Caption         =   "����Ͻ�"
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   465
      Left            =   1350
      TabIndex        =   34
      Top             =   10125
      Width           =   1305
   End
   Begin VB.Label lbl_Name 
      BackStyle       =   0  '����
      Caption         =   "�� �� ��"
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   465
      Left            =   1350
      TabIndex        =   33
      Top             =   10590
      Width           =   1305
   End
   Begin VB.Label lbl_CarNo 
      BackStyle       =   0  '����
      Caption         =   "�Ǹż���"
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   465
      Left            =   5580
      TabIndex        =   32
      Top             =   11100
      Width           =   1305
   End
   Begin VB.Label lbl_title 
      BackStyle       =   0  '����
      Caption         =   "# ���α� ��� / ����"
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   465
      Index           =   1
      Left            =   570
      TabIndex        =   31
      Top             =   9510
      Width           =   4395
   End
   Begin VB.Label lbl_Search 
      Alignment       =   1  '������ ����
      BackStyle       =   0  '����
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   15.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   15570
      TabIndex        =   28
      Top             =   3240
      Width           =   2025
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '����
      Caption         =   "���α� �Ǹ� ����"
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
      Left            =   1470
      TabIndex        =   27
      Top             =   870
      Width           =   3525
   End
   Begin VB.Label lbl_title 
      BackStyle       =   0  '����
      Caption         =   "# ���� ���� ��Ȳ"
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   14.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Index           =   0
      Left            =   480
      TabIndex        =   26
      Top             =   3180
      Width           =   2895
   End
   Begin VB.Label lbl_COUNT 
      BackStyle       =   0  '����
      Caption         =   "0000"
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
      Height          =   375
      Left            =   11850
      TabIndex        =   25
      Top             =   3300
      Width           =   1425
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '����
      Caption         =   "��ȸ�� �����Ǽ� :"
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
      Height          =   375
      Index           =   0
      Left            =   9780
      TabIndex        =   24
      Top             =   3300
      Width           =   1875
   End
   Begin VB.Label lbl_title 
      BackStyle       =   0  '����
      Caption         =   "������ �˻�"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   20.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   495
      Index           =   2
      Left            =   600
      TabIndex        =   23
      Top             =   2100
      Width           =   2115
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '����
      Caption         =   "��  ��  �� :"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Index           =   2
      Left            =   3420
      TabIndex        =   22
      Top             =   2220
      Width           =   1305
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "��ȸ�Ⱓ :"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   5
      Left            =   7170
      TabIndex        =   21
      Top             =   2205
      Width           =   1035
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
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   7
      Left            =   10770
      TabIndex        =   20
      Top             =   2205
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
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   9
      Left            =   13980
      TabIndex        =   19
      Top             =   2205
      Width           =   450
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "Total : "
      BeginProperty Font 
         Name            =   "�������"
         Size            =   14.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   315
      Index           =   0
      Left            =   14460
      TabIndex        =   18
      Top             =   3270
      Width           =   900
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H006F3C2F&
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00C0C0C0&
      Height          =   1095
      Left            =   270
      Top             =   1800
      Width           =   18675
   End
End
Attribute VB_Name = "FrmCouponSale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CAR_NO_TMP As String
Dim INDEX_NO_TMP As String
Dim PART_NAME_TMP As String
Dim CN1, CN2, CN3, CN4, CN5, CN6, CN7 As String
Dim CS1, CS2, CS3, CS4, CS5, CS6, CS7 As Long

Private Sub Form_Load()
Dim i, s As Integer

'Left = (Screen.Width - Width) / 2   ' ���� ���η� �߾ӿ� �����ϴ�.
'Top = (Screen.Height - Height) / 2   ' ���� ���η� �߾ӿ� �����ϴ�.
Left = 0
Top = 0

cmb_Sort.ListIndex = 0

DTPicker1.value = Now
DTPicker2.value = Now

'INI ����
'[���α� ����]
'CN1=1�ð��湮Ȯ��
'CS1 = 100
'CN2=2�ð��湮Ȯ��
'CS2 = 200
'CN3=3�ð��湮Ȯ��
'CS3 = 300
'CN4=1�ð��ʰ����α�
'CS4 = 400
'CN5=2�ð��ʰ����α�
'CS5 = 500
'CN6=3�ð��ʰ����α�
'CS6 = 600
'CN7 = �������α�
'CS7 = 1000

'���α� ����
i = Val(Get_Ini("���α� ����", "CouponType", ""))
CN1 = Get_Ini("���α� ����", "CN1", "")
CN2 = Get_Ini("���α� ����", "CN2", "")
CN3 = Get_Ini("���α� ����", "CN3", "")
CN4 = Get_Ini("���α� ����", "CN4", "")
CN5 = Get_Ini("���α� ����", "CN5", "")
CN6 = Get_Ini("���α� ����", "CN6", "")
CN7 = Get_Ini("���α� ����", "CN7", "")
CS1 = Val(Get_Ini("���α� ����", "CS1", ""))
CS2 = Val(Get_Ini("���α� ����", "CS2", ""))
CS3 = Val(Get_Ini("���α� ����", "CS3", ""))
CS4 = Val(Get_Ini("���α� ����", "CS4", ""))
CS5 = Val(Get_Ini("���α� ����", "CS5", ""))
CS6 = Val(Get_Ini("���α� ����", "CS6", ""))
CS7 = Val(Get_Ini("���α� ����", "CS7", ""))
With cmb_Gubun
    .AddItem CN1
    .AddItem CN2
    .AddItem CN3
    .AddItem CN4
    .AddItem CN5
    .AddItem CN6
    .AddItem CN7
    .Text = cmb_Gubun.List(0)
End With
MaskEdBox_Fee(0) = CS1

Glo_SQL_REG = "SELECT * FROM TB_COUPON_SALE WHERE (REG_DATE >= '" & Format(DTPicker1, "yyyy-mm-dd") & " 00:00:00') AND (REG_DATE <= '" & Format(DTPicker2, "yyyy-mm-dd") & " 23:59:59') ORDER BY REG_DATE ASC"
Call Clear_Field
Call ListView_REG_Draw
Call ListView_REG_SQL

List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    ���α� ���/���� ����...!!", 0
Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    ���α� ���/���� ����...!!")

End Sub

Public Sub ListView_REG_SQL()
Dim rs As Recordset
Dim QRY As String
Dim itmX As ListItem
Dim INDEX_NO As Long

lbl_Search = 0

INDEX_NO = 1
Set rs = New ADODB.Recordset
rs.Open Glo_SQL_REG, adoConn
lbl_COUNT = rs.RecordCount

Do While Not (rs.EOF)
    Set itmX = ListView_REG.ListItems.Add(, , "" & INDEX_NO)
    itmX.SubItems(1) = "" & rs!SALE_DATE
    itmX.SubItems(2) = "" & rs!CONSUMER
    itmX.SubItems(3) = "" & rs!CONSUMER_PHONE
    itmX.SubItems(4) = "" & rs!COUPON_NAME
    itmX.SubItems(5) = "" & rs!COUPON_START & " ~ " & rs!COUPON_END
    itmX.SubItems(6) = "" & rs!COUPON_PRICE
    itmX.SubItems(7) = "" & rs!COUPON_NUM
    itmX.SubItems(8) = "" & rs!SALE_AMOUNT
    itmX.SubItems(9) = "" & rs!REG_DATE
    itmX.SubItems(10) = "" & rs!Update_date
    lbl_Search = lbl_Search + rs!SALE_AMOUNT
    rs.MoveNext
    INDEX_NO = INDEX_NO + 1
Loop
Set rs = Nothing
End Sub

Public Sub ListView_REG_Draw()
Dim Column_to_size As Integer

With Me
    Call ListViewExtended(.ListView_REG)
    .ListView_REG.View = lvwReport
    .ListView_REG.ListItems.Clear
    .ListView_REG.ColumnHeaders.Clear
    .ListView_REG.ColumnHeaders.Add , , " No  "
    .ListView_REG.ColumnHeaders.Add , , " �Ǹ�����       "
    .ListView_REG.ColumnHeaders.Add , , " �� �� ��       "
    .ListView_REG.ColumnHeaders.Add , , " �� �� ó                 "
    .ListView_REG.ColumnHeaders.Add , , " �� �� ��                 "
    .ListView_REG.ColumnHeaders.Add , , " �� �� �� �� ȣ   "
    .ListView_REG.ColumnHeaders.Add , , " ��  ��   "
    .ListView_REG.ColumnHeaders.Add , , " �Ǹż���    "
    .ListView_REG.ColumnHeaders.Add , , " �Ǹűݾ�      "
    .ListView_REG.ColumnHeaders.Add , , " �� �� ��                 "
    .ListView_REG.ColumnHeaders.Add , , " �� �� ��                 "
    .ListView_REG.ColumnHeaders.Add , , " "
    For Column_to_size = 0 To .ListView_REG.ColumnHeaders.Count - 2
         SendMessage .ListView_REG.hwnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
    Next
End With
End Sub

Private Sub ListView_REG_ItemClick(ByVal Item As ComctlLib.ListItem)
Dim i, j As Integer
Dim str As String

ListView_REG.SetFocus
txt_Num = ListView_REG.SelectedItem.SubItems(9)
txt_Name = ListView_REG.SelectedItem.SubItems(2)
txt_Phone = ListView_REG.SelectedItem.SubItems(3)
cmb_Gubun = ListView_REG.SelectedItem.SubItems(4)
'MaskEdBox_Fee(0) = ListView_REG.SelectedItem.SubItems(5)
MaskEdBox_Fee(2) = ListView_REG.SelectedItem.SubItems(7)
MaskEdBox_Fee(1) = ListView_REG.SelectedItem.SubItems(8)
j = Len(ListView_REG.SelectedItem.SubItems(5))
i = InStr(ListView_REG.SelectedItem.SubItems(5), " ")
str = Left(ListView_REG.SelectedItem.SubItems(5), i - 1)
txt_Dong = str
str = Right(ListView_REG.SelectedItem.SubItems(5), j - i - 2)
txt_Ho = str

End Sub

Public Sub Clear_Field()
txt_tmpName = ""
'DTPicker1.value = Now
'DTPicker2.value = Now

CAR_NO_TMP = ""
INDEX_NO_TMP = ""
txt_Num.Text = ""
txt_Name.Text = ""
txt_Phone.Text = ""
cmb_Gubun.ListIndex = 0
txt_Dong.Text = ""
txt_Ho.Text = ""
'MaskEdBox_Fee(0).Text = "0"
MaskEdBox_Fee(1).Text = "0"
MaskEdBox_Fee(2).Text = "0"

On Error Resume Next
txt_Name.SetFocus
End Sub

'������ ����
Sub Delete_Record()
adoConn.Execute "DELETE FROM TB_COUPON_SALE WHERE REG_DATE = '" & txt_Num & "'"
List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_Name & "    ���α� ���� ���� �Ϸ�", 0
Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_Name & "    ���α� ���� ���� �Ϸ�")
Call ListView_REG_Draw
Call ListView_REG_SQL
End Sub

Sub Insert_Record()
Dim rs_COUNT As Recordset
Dim rs As Recordset
Dim SQL_COUNT As String
Dim SQL_QUARY As String
Dim i As Integer
Dim Cnt As Integer
Dim tmp As String
Dim P As String

If (txt_Num = "") Then '�űԵ��
    'INSERT
    adoConn.Execute "INSERT INTO TB_COUPON_SALE VALUES ('" & Format(Now, "YYYY-MM-DD") & "', '" & txt_Name & "', '" & txt_Phone & "', '" & cmb_Gubun & "', '" & txt_Dong & "', '" & txt_Ho & "', '" & MaskEdBox_Fee(2) & "', '" & MaskEdBox_Fee(0) & "', '" & MaskEdBox_Fee(1) & "', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "', '')"
    List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_Name & " " & cmb_Gubun & " " & MaskEdBox_Fee(1) & "��    ���α� ��� �Ϸ�", 0
    Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_Name & " " & cmb_Gubun & " " & MaskEdBox_Fee(1) & "��    ���α� ��� �Ϸ�")
Else
    adoConn.Execute "UPDATE TB_COUPON_SALE SET CONSUMER = '" & txt_Name & "', CONSUMER_PHONE = '" & txt_Phone & "', COUPON_NAME = '" & cmb_Gubun & "', COUPON_START = '" & txt_Dong & "', COUPON_END = '" & txt_Ho & "', COUPON_NUM = '" & MaskEdBox_Fee(2) & "', COUPON_PRICE = '" & MaskEdBox_Fee(0) & "', SALE_AMOUNT = '" & MaskEdBox_Fee(1) & "', UPDATE_DATE = '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "' WHERE REG_DATE = '" & txt_Num & "'"
    List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_Name & "    ���α� ���� �Ϸ�", 0
    Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_Name & "    ���α� ���� �Ϸ�")
End If

Call ListView_REG_Draw
Call ListView_REG_SQL

On Error Resume Next
    If (Err = 3022) Then
        Msg_Box.Label2.Caption = "������ ���̽� ����"
        Msg_Box.Label1.Caption = "�ߺ��� ������ȣ�� ��������ʽ��ϴ�."
        Msg_Box.Show 1
    End If

End Sub

Private Sub cmb_Gubun_Click()
Select Case cmb_Gubun.ListIndex
    Case 0
        MaskEdBox_Fee(0) = CS1
    Case 1
        MaskEdBox_Fee(0) = CS2
    Case 2
        MaskEdBox_Fee(0) = CS3
    Case 3
        MaskEdBox_Fee(0) = CS4
    Case 4
        MaskEdBox_Fee(0) = CS5
    Case 5
        MaskEdBox_Fee(0) = CS6
    Case 6
        MaskEdBox_Fee(0) = CS7
End Select
End Sub

Private Sub cmd_Button_Click(Index As Integer)
Dim i, j As Integer
Dim myExcelFile As New ExcelFile
Dim tmpFileName As String
Dim str As String


Select Case Index
    Case 0  '����
        List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    �������/���� ����", 0
        Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    �������/���� ����")
        Unload Me
        Exit Sub
    
    Case 1
        tmpFileName = Format(Now, "YYYYMMDD_HHMMSS")
        tmpFileName = App.Path & "\Excel\" & tmpFileName & "_���αǳ���" & ".xls"
        Call makeexcel(ListView_REG, tmpFileName, "���α�")
        Exit Sub
    
    Case 2
        Select Case cmb_Sort.ListIndex
            Case 0
                str = "CONSUMER"
            Case 1
                str = "COUPON_NAME"
            Case 2
                str = "REG_DATE"
        End Select
        
        
        If (Len(txt_tmpName) = 0) Then
            Glo_SQL_REG = "SELECT * FROM TB_COUPON_SALE WHERE (REG_DATE >= '" & Format(DTPicker1, "yyyy-mm-dd") & " 00:00:00') AND (REG_DATE <= '" & Format(DTPicker2, "yyyy-mm-dd") & " 23:59:59') ORDER BY " & str & ""
        Else
            Glo_SQL_REG = "SELECT * FROM TB_COUPON_SALE WHERE (REG_DATE >= '" & Format(DTPicker1, "yyyy-mm-dd") & " 00:00:00') AND (REG_DATE <= '" & Format(DTPicker2, "yyyy-mm-dd") & " 23:59:59') AND CONSUMER = '" & txt_tmpName & "' ORDER BY " & str & ""
        End If
        Call Clear_Field
        Call ListView_REG_Draw
        Call ListView_REG_SQL
        Exit Sub
    
    Case 3  '����
        If (txt_Num = "") Then
           Call Clear_Field
           Exit Sub
        End If
        MBox.Label3.Caption = txt_Name.Text
        MBox.Label1.Caption = "�� �ŷ�ó�� ���α� ��� ������ �����մϴ�." & vbCrLf & vbCrLf & " �����Ͻðڽ��ϱ�?"
        MBox.Label2.Caption = "���α� ��� ���� ����"
        MBox.Show 1
        If (Glo_MsgRet = True) Then
           Call Delete_Record
        End If
        Call Clear_Field
        Exit Sub
            
    Case 4  '����
        If (txt_Num = "") Then
            Msg_Box.Label2.Caption = "�ʵ� ����"
            Msg_Box.Label1.Caption = "�ű� ����ڷ� �Դϴ�." & vbCrLf & vbCrLf & " �ٽ� Ȯ�� �ϼ���."
            Msg_Box.Show 1
            Exit Sub
        Else
            If (Data_Error_Check = False) Then
                Msg_Box.Label2.Caption = "�ʵ� �Է� ����"
                Msg_Box.Label1.Caption = "�߿��� �׸��� ���� �Ǵ� �߸� �Է��Ͽ����ϴ�."
                Msg_Box.Show 1
            Else
                MBox.Label3.Caption = txt_Name.Text
                MBox.Label1.Caption = "�����Ͻ� ���α� ��� ������ ����˴ϴ�." & vbCrLf & vbCrLf & " ���� �Ͻðڽ��ϱ�?"
                MBox.Label2.Caption = "���α� �ڷ� ����"
                MBox.Show 1
                If (Glo_MsgRet = True) Then
                   Call Insert_Record
                   Call Clear_Field
                   'txt_CarNo.SetFocus
                End If
            End If
            
        End If
        Exit Sub

    Case 5  '�ű��Է�
        If (txt_Num = "") Then
            If (Data_Error_Check = False) Then
                Msg_Box.Label2.Caption = "�ʵ� �Է� ����"
                Msg_Box.Label1.Caption = "�߿��� �׸��� �Է����� �ʾҽ��ϴ�."
                Msg_Box.Show 1
            Else
                Call Insert_Record
                Call Clear_Field
            End If
        Else
            Msg_Box.Label2.Caption = "�ű� ������ �Է� ����"
            Msg_Box.Label1.Caption = "�ű� �����Ͱ� �ƴմϴ�." & vbCrLf & vbCrLf & " �ٽ� �ѹ� Ȯ���ϼ���."
            Msg_Box.Show 1
            'Call Clear_Field
        End If
        Exit Sub
     
     Case 6   '�ʱ�ȭ
        Call Clear_Field
        Exit Sub
End Select

On Error Resume Next

End Sub


'�ʼ� �Է� ������ Ȯ��
Private Function Data_Error_Check()
Dim Error_Flag As Boolean
    
Error_Flag = True

If (Len(txt_Name.Text) = 0) Then
    'txt_Name.Text = ""
    Error_Flag = False
Else
txt_Name.Text = MidH(txt_Name.Text, 1, 16)

End If
If (LenH(txt_Phone.Text) = 0) Then
    'txt_Phone.Text = " "
    'Error_Flag = False
Else
    txt_Phone.Text = Mid(txt_Phone.Text, 1, 16)
End If

If (IsNumeric(MaskEdBox_Fee(0).Text) = False) Then
    Error_Flag = False
End If
If (IsNumeric(MaskEdBox_Fee(1).Text) = False) Then
    Error_Flag = False
End If
If (IsNumeric(MaskEdBox_Fee(2).Text) = False) Then
    Error_Flag = False
End If
If (txt_Ho <> "") Then
    If IsNumeric(txt_Ho.Text) Then
        'txt_Ho.Text = Format(txt_Ho.Text, "0000")
    Else
        MsgBox "�����ȣ �Է��� Ȯ���ϼ���...!!"
    End If
End If
If (txt_Dong <> "") Then
    If IsNumeric(txt_Dong.Text) Then
        'txt_Dong.Text = Format(txt_Dong.Text, "0000")
    Else
        MsgBox "�����ȣ �Է��� Ȯ���ϼ���...!!"
    End If
End If

Data_Error_Check = Error_Flag
End Function

'����Ű �Է½� �� ����
'���Ӽ� keypreview = true ����
Private Sub Form_KeyPRESS(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    SendKeys "{TAB}"
End If
End Sub

Private Sub MaskEdBox_Fee_Change(Index As Integer)
'MaskEdBox_Fee(1) = ""
If Index = 2 Then
    If (Val(MaskEdBox_Fee(0)) <> 0) And (Val(MaskEdBox_Fee(2)) <> 0) Then
        MaskEdBox_Fee(1) = Val(MaskEdBox_Fee(0)) * Val(MaskEdBox_Fee(2))
    End If
End If
End Sub
