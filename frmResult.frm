VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMctl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmResult 
   BorderStyle     =   1  '���� ����
   Caption         =   "��������"
   ClientHeight    =   14535
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   19200
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmResult.frx":0000
   ScaleHeight     =   14535
   ScaleWidth      =   19200
   Begin VB.ComboBox cmb_Partner 
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
      Left            =   12255
      TabIndex        =   40
      Top             =   1005
      Width           =   2205
   End
   Begin VB.ComboBox cmb_Search 
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
      Left            =   12255
      TabIndex        =   39
      Top             =   1950
      Width           =   2205
   End
   Begin VB.ComboBox cmb_DCGubun 
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
      Left            =   12255
      TabIndex        =   36
      Top             =   2880
      Width           =   2205
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   13680
      Width           =   18975
   End
   Begin VB.TextBox txt_CarNo 
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
      Left            =   12255
      TabIndex        =   1
      Top             =   2385
      Width           =   2205
   End
   Begin VB.TextBox txt_Name 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�������"
         Size            =   14.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   435
      Left            =   21780
      TabIndex        =   0
      Top             =   12270
      Visible         =   0   'False
      Width           =   2115
   End
   Begin ComctlLib.ListView ListView_REG 
      Height          =   4635
      Left            =   255
      TabIndex        =   3
      Top             =   3600
      Width           =   18690
      _ExtentX        =   32967
      _ExtentY        =   8176
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
      Left            =   17355
      TabIndex        =   4
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
      Picture         =   "frmResult.frx":F4A3
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   660
      Index           =   1
      Left            =   17355
      TabIndex        =   5
      Top             =   1980
      Width           =   1410
      _Version        =   65536
      _ExtentX        =   2487
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
      Picture         =   "frmResult.frx":F7F4
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   660
      Index           =   2
      Left            =   14805
      TabIndex        =   6
      Top             =   1980
      Width           =   1410
      _Version        =   65536
      _ExtentX        =   2487
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
      Picture         =   "frmResult.frx":FB45
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   5220
      TabIndex        =   7
      Top             =   2130
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   661
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
      Format          =   138149888
      CurrentDate     =   36927
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   5220
      TabIndex        =   8
      Top             =   2760
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   661
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
      Format          =   138149888
      CurrentDate     =   36927
   End
   Begin MSComCtl2.DTPicker DTPicker3 
      Height          =   345
      Left            =   7440
      TabIndex        =   9
      Top             =   2130
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
      Format          =   138149890
      CurrentDate     =   36927
   End
   Begin MSComCtl2.DTPicker DTPicker4 
      Height          =   345
      Left            =   7425
      TabIndex        =   10
      Top             =   2775
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
      Format          =   138149890
      CurrentDate     =   36927
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   660
      Index           =   3
      Left            =   14745
      TabIndex        =   41
      Top             =   690
      Width           =   1410
      _Version        =   65536
      _ExtentX        =   2487
      _ExtentY        =   1164
      _StockProps     =   78
      Caption         =   "������ �˻�"
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
      Picture         =   "frmResult.frx":FE96
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   660
      Index           =   4
      Left            =   15045
      TabIndex        =   42
      Top             =   12915
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   1164
      _StockProps     =   78
      Caption         =   "ī����� ���"
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
      RoundedCorners  =   0   'False
      Picture         =   "frmResult.frx":101E7
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   660
      Index           =   5
      Left            =   12555
      TabIndex        =   43
      Top             =   12915
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   1164
      _StockProps     =   78
      Caption         =   "������ �����"
      ForeColor       =   8454143
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
      Picture         =   "frmResult.frx":10538
   End
   Begin MSWinsockLib.Winsock RePrint_Sock 
      Left            =   12075
      Top             =   13140
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock CardCancel_Sock 
      Left            =   14550
      Top             =   13140
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '����
      Caption         =   "�� �� :"
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
      Index           =   3
      Left            =   2295
      TabIndex        =   56
      Top             =   11685
      Width           =   915
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '����
      Caption         =   "�ſ�ī�� :"
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
      Index           =   4
      Left            =   1875
      TabIndex        =   55
      Top             =   12165
      Width           =   1335
   End
   Begin VB.Label lbl_Cash 
      BackStyle       =   0  '����
      Caption         =   "0"
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
      Left            =   3525
      TabIndex        =   54
      Top             =   11685
      Width           =   1905
   End
   Begin VB.Label lbl_Card 
      BackStyle       =   0  '����
      Caption         =   "0"
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
      Left            =   3525
      TabIndex        =   53
      Top             =   12165
      Width           =   1905
   End
   Begin VB.Label lbl_TotalSum 
      BackStyle       =   0  '����
      Caption         =   "0"
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
      Left            =   3525
      TabIndex        =   52
      Top             =   10125
      Width           =   1905
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '����
      Caption         =   "������� :"
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
      Index           =   5
      Left            =   1905
      TabIndex        =   51
      Top             =   10125
      Width           =   1245
   End
   Begin VB.Label lbl_DcSum 
      BackStyle       =   0  '����
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   3525
      TabIndex        =   50
      Top             =   10605
      Width           =   1905
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '����
      Caption         =   "���ο�� :"
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
      Index           =   6
      Left            =   1905
      TabIndex        =   49
      Top             =   10605
      Width           =   1245
   End
   Begin VB.Label lbl_RealSum 
      BackStyle       =   0  '����
      Caption         =   "0"
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
      Height          =   375
      Left            =   3525
      TabIndex        =   48
      Top             =   11205
      Width           =   1905
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '����
      Caption         =   "���Ϳ�� :"
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
      Index           =   7
      Left            =   1905
      TabIndex        =   47
      Top             =   11205
      Width           =   1245
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      X1              =   1095
      X2              =   7035
      Y1              =   11055
      Y2              =   11055
   End
   Begin VB.Label lbl_seq 
      Caption         =   "Label1"
      Height          =   240
      Left            =   19470
      TabIndex        =   46
      Top             =   8100
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label lbl_ticketcode 
      Caption         =   "lbl_ticketcode"
      Height          =   240
      Left            =   19470
      TabIndex        =   45
      Top             =   7470
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label lbl_carno 
      Caption         =   "lbl_carno"
      Height          =   240
      Left            =   19470
      TabIndex        =   44
      Top             =   7770
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�˻����� :"
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
      Index           =   3
      Left            =   11025
      TabIndex        =   38
      Top             =   1950
      Width           =   1035
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�����׸� :"
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
      Index           =   2
      Left            =   11025
      TabIndex        =   37
      Top             =   2895
      Width           =   1035
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
      Left            =   15210
      TabIndex        =   35
      Top             =   8520
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
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   9
      Left            =   9540
      TabIndex        =   34
      Top             =   2820
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
      Index           =   7
      Left            =   9540
      TabIndex        =   33
      Top             =   2175
      Width           =   450
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
      Left            =   4110
      TabIndex        =   32
      Top             =   2175
      Width           =   1035
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
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   2
      Left            =   510
      TabIndex        =   31
      Top             =   2010
      Width           =   2115
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
      Left            =   10530
      TabIndex        =   30
      Top             =   8550
      Width           =   1875
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
      Left            =   12600
      TabIndex        =   29
      Top             =   8550
      Width           =   1425
   End
   Begin VB.Label lbl_title 
      BackStyle       =   0  '����
      Caption         =   "# ���� ����"
      BeginProperty Font 
         Name            =   "���� ���"
         Size            =   20.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   615
      Index           =   0
      Left            =   12300
      TabIndex        =   28
      Top             =   9480
      Width           =   3060
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '����
      Caption         =   "���� ���� "
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
   Begin VB.Label lbl_Search 
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
      Left            =   16320
      TabIndex        =   26
      Top             =   8490
      Width           =   2025
   End
   Begin VB.Label lbl_Name 
      BackStyle       =   0  '����
      Caption         =   "�����Ͻ� :"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   15.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   435
      Index           =   0
      Left            =   12510
      TabIndex        =   25
      Top             =   10170
      Width           =   1665
   End
   Begin VB.Label lbl_Name 
      BackStyle       =   0  '����
      Caption         =   "�����Ͻ� :"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   15.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   435
      Index           =   1
      Left            =   12510
      TabIndex        =   24
      Top             =   10695
      Width           =   1665
   End
   Begin VB.Label lbl_Name 
      BackStyle       =   0  '����
      Caption         =   "������� :"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   15.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   435
      Index           =   2
      Left            =   12510
      TabIndex        =   23
      Top             =   12270
      Width           =   1665
   End
   Begin VB.Label lbl_Name 
      BackStyle       =   0  '����
      Caption         =   "�������ð� :"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   15.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   435
      Index           =   3
      Left            =   12510
      TabIndex        =   22
      Top             =   11220
      Width           =   1710
   End
   Begin VB.Label lbl_Name 
      BackStyle       =   0  '����
      Caption         =   "POS CODE :"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "�������"
         Size            =   15.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   435
      Index           =   4
      Left            =   12510
      TabIndex        =   21
      Top             =   12795
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lbl_Name 
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "�������"
         Size            =   15.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   435
      Index           =   5
      Left            =   14280
      TabIndex        =   20
      Top             =   10170
      Width           =   4545
   End
   Begin VB.Label lbl_Name 
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "�������"
         Size            =   15.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   435
      Index           =   6
      Left            =   14280
      TabIndex        =   19
      Top             =   10695
      Width           =   4545
   End
   Begin VB.Label lbl_Name 
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "�������"
         Size            =   15.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   435
      Index           =   7
      Left            =   14280
      TabIndex        =   18
      Top             =   12270
      Width           =   4545
   End
   Begin VB.Label lbl_Name 
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "�������"
         Size            =   15.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   435
      Index           =   8
      Left            =   14280
      TabIndex        =   17
      Top             =   11220
      Width           =   4545
   End
   Begin VB.Label lbl_Name 
      BackStyle       =   0  '����
      Enabled         =   0   'False
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
      Height          =   435
      Index           =   9
      Left            =   14280
      TabIndex        =   16
      Top             =   12795
      Visible         =   0   'False
      Width           =   4545
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "������ȣ :"
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
      Index           =   1
      Left            =   11025
      TabIndex        =   15
      Top             =   2430
      Width           =   1035
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '����
      Caption         =   "������ȣ :"
      Enabled         =   0   'False
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
      Height          =   465
      Index           =   1
      Left            =   20280
      TabIndex        =   14
      Top             =   11730
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '����
      Caption         =   "��       �� :"
      Enabled         =   0   'False
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
      Height          =   465
      Index           =   2
      Left            =   20280
      TabIndex        =   13
      Top             =   12330
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label lbl_Name 
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "�������"
         Size            =   15.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   435
      Index           =   10
      Left            =   14280
      TabIndex        =   12
      Top             =   11745
      Width           =   4545
   End
   Begin VB.Label lbl_Name 
      BackStyle       =   0  '����
      Caption         =   "���νð� :"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   15.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   435
      Index           =   11
      Left            =   12510
      TabIndex        =   11
      Top             =   11745
      Width           =   1665
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H006F3C2F&
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00C0C0C0&
      Height          =   1620
      Left            =   255
      Top             =   1770
      Width           =   18690
   End
End
Attribute VB_Name = "frmResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bShow_Partner As Boolean
Dim sSelButton As String '�����ΰ˻�, �˻�


Private Sub Form_Load()
    Dim i As Integer
    Dim rs As Recordset
    Dim qry As String
    Dim sStartDT As String
    Dim sEndDT As String
    'cmb_DCGubun
    qry = "SELECT tb_calculate.DC_GUBUN From tb_calculate Group By tb_calculate.DC_GUBUN "
    Set rs = New ADODB.Recordset
    rs.Open qry, adoConn
    
    With cmb_DCGubun
        .AddItem "��ü"
        Do While Not (rs.EOF)
            If (rs!DC_Gubun = "0") Then
                .AddItem "������"
            Else
                .AddItem rs!DC_Gubun
            End If
            rs.MoveNext
        Loop
        Set rs = Nothing
        .text = cmb_DCGubun.List(0)
    End With
    
    Left = (Screen.width - width) / 2   ' ���� ���η� �߾ӿ� �����ϴ�.
    Top = (Screen.height - height) / 2   ' ���� ���η� �߾ӿ� �����ϴ�.
    'Left = 0
    'Top = 0
    'Me.cmb_Gubun = Me.cmb_Gubun.List(0)
    
    
    '������ ��Ʈ�� �����ֱ�
    Call Load_WebDC_Partner
    
    With cmb_Search
        .AddItem "��ü"
        .AddItem "�ſ�ī��"
        .text = cmb_Search.List(0)
    End With
        
    DTPicker1.value = Now
    DTPicker2.value = Now
    
    DTPicker3.Format = dtpCustom
    DTPicker3.CustomFormat = "HH:mm:ss"
    DTPicker3.Refresh
    
    DTPicker4.Format = dtpCustom
    DTPicker4.CustomFormat = "HH:mm:ss"
    DTPicker4.Refresh
    
    DTPicker3.value = Format("00:00:00")
    DTPicker4.value = Format("23:59:59")
    
    
    sStartDT = Format(DTPicker1, "yyyy-mm-dd") & " " & Format(DTPicker3, "hh:nn:ss")
    sEndDT = Format(DTPicker2, "yyyy-mm-dd") & " " & Format(DTPicker4, "hh:nn:ss")
    
'''    sSelButton = "�� ��"
'''    Glo_SQL_REG = "SELECT * FROM tb_calculate WHERE ( REG_DATE >= '" & sStartDT & "') AND ( REG_DATE <= '" & sEndDT & "') ORDER BY REG_DATE"
'''    'Glo_SQL_REG = "SELECT * From TB_FEE WHERE ORDER BY REG_DATE ASC"
'''    Call ListView_REG_Draw
'''    Call ListView_REG_SQL
'''    Call ListView_Reg_ClearCaption
'''    List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    �������� ����...!!", 0
'''    Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    �������� ����...!!")
    'Call cmd_Button_Click(3)
    Call cmd_Button_Click(2)
End Sub

Private Sub Load_WebDC_Partner()
    Dim rs As Recordset
    Dim rs2 As Recordset
    
    Set rs = New ADODB.Recordset
    'rs.Open "SELECT ID FROM tb_id WHERE GUBUN='��Ʈ��' ", adoConn
    rs.Open "SELECT ID FROM tb_id WHERE MENU8='������' ", adoConn
    
    If (Not (rs.EOF)) Then
        bShow_Partner = True
        cmb_Partner.Enabled = True
        cmb_Partner.Visible = True
        cmb_Partner.AddItem "��ü"
        cmd_Button(3).Enabled = True
        cmd_Button(3).Visible = True
        
        Do While Not (rs.EOF)
            Set rs2 = New ADODB.Recordset
            rs2.Open "SELECT PNAME FROM tb_partner WHERE ID='" & rs!ID & "' ", adoConn
            If (Not (rs2.EOF)) Then
                cmb_Partner.AddItem "" & rs2!PNAME
            End If
            Set rs2 = Nothing
            
            rs.MoveNext
        Loop
        cmb_Partner.ListIndex = 0
    Else
        bShow_Partner = False
        cmb_Partner.Enabled = False
        cmb_Partner.Visible = False
        cmd_Button(3).Enabled = False
        cmd_Button(3).Visible = False
        cmd_Button(3).value = False
    End If
    Set rs = Nothing
End Sub

Private Sub ListView_Reg_ClearCaption()
    lbl_title(0).Caption = "# ���� ����"
    lbl_Name(0).Caption = "�����Ͻ�:"
    lbl_Name(1).Caption = "�����Ͻ�:"
    lbl_Name(3).Caption = "�������ð�:"
    lbl_Name(11).Caption = "���νð�:"
    lbl_Name(2).Caption = "�������:"
    lbl_Name(4).Caption = "POS CODE :"
    
    lbl_Name(5).Caption = ""
    lbl_Name(6).Caption = ""
    lbl_Name(8).Caption = ""
    lbl_Name(10).Caption = ""
    lbl_Name(7).Caption = ""
    lbl_Name(9).Caption = ""
End Sub

Public Sub ListView_REG_SQL()
    Dim rs As Recordset
    Dim rs2 As Recordset
    Dim qry As String
    Dim itmX As ListItem
    Dim INDEX_NO As Long
    Dim TOTAL_FEE As Single
    Dim SumCash As Long
    Dim SumCard As Long
    Dim DcSum As Long
    Dim TotalSum As Long
    Dim Tmp_DCNames As String
    Dim Col_Idx As Long
    

    INDEX_NO = 1
    TOTAL_FEE = 0
    Set rs = New ADODB.Recordset
    rs.Open Glo_SQL_REG, adoConn
    lbl_COUNT = rs.RecordCount
    Do While Not (rs.EOF)
    
        'Set itmX = ListView_REG.ListItems.Add(, , "" & INDEX_NO)
        Set itmX = ListView_REG.ListItems.Add(, , "" & rs!SEQ)
    
        Col_Idx = 1
        itmX.SubItems(Col_Idx) = "" & rs!TICKET_CODE: Col_Idx = Col_Idx + 1
        itmX.SubItems(Col_Idx) = "" & rs!IN_DATE: Col_Idx = Col_Idx + 1
        itmX.SubItems(Col_Idx) = "" & rs!IN_TIME: Col_Idx = Col_Idx + 1
        itmX.SubItems(Col_Idx) = "" & rs!OUT_DATE: Col_Idx = Col_Idx + 1
        itmX.SubItems(Col_Idx) = "" & rs!OUT_TIME: Col_Idx = Col_Idx + 1
        itmX.SubItems(Col_Idx) = "" & rs!TICKET_NO: Col_Idx = Col_Idx + 1           '������ȣ
        
        
        'itmX.SubItems(Col_Idx) = "" & rs!TOTAL_PARKING_TIME: Col_Idx = Col_Idx + 1  '�����ð�
        If (rs!TOTAL_PARKING_TIME >= 60) Then
            itmX.SubItems(Col_Idx) = "" & Int(rs!TOTAL_PARKING_TIME / 60) & "�ð� "
            
            'Debug.Print "->" & rs!TOTAL_PARKING_TIME
            
            If (rs!TOTAL_PARKING_TIME Mod 60) Then
                itmX.SubItems(Col_Idx) = itmX.SubItems(Col_Idx) & rs!TOTAL_PARKING_TIME Mod 60 & "��"
            End If
        Else
            itmX.SubItems(Col_Idx) = Val("" & rs!TOTAL_PARKING_TIME) & "�� "
        End If
        Col_Idx = Col_Idx + 1
        
        
        itmX.SubItems(Col_Idx) = "" & rs!DC_TIME: Col_Idx = Col_Idx + 1                                         '���νð�
        itmX.SubItems(Col_Idx) = "" & rs!DC_MONEY: Col_Idx = Col_Idx + 1                                        '���αݾ�
        itmX.SubItems(Col_Idx) = "" & Format(rs!TOTAL_PARKING_PAYMENT, "#,###,##0��"): Col_Idx = Col_Idx + 1    '�������
        itmX.SubItems(Col_Idx) = "" & rs!Gubun: Col_Idx = Col_Idx + 1                                           '����
        itmX.SubItems(Col_Idx) = "" & Format(rs!TOTAL_PAID, "#,###,##0��"): Col_Idx = Col_Idx + 1               '�����ݾ�

            Tmp_DCNames = ""
            qry = "SELECT M_NAME, DC_CODE From tb_dc_log WHERE RECEIPT_NO = '" & rs!TICKET_CODE & "' AND DT_DATE = '" & Format(rs!REG_DATE, "yyyy-mm-dd hh:nn:ss") & "' "
            Set rs2 = New ADODB.Recordset
            rs2.Open qry, adoConn
            Do While Not (rs2.EOF)
                Tmp_DCNames = Tmp_DCNames & rs2!M_NAME & "(" & rs2!DC_CODE & ")" & ","
                rs2.MoveNext
            Loop
            Set rs2 = Nothing
            If (Right(Tmp_DCNames, 1) = ",") Then
                Tmp_DCNames = LeftH(Tmp_DCNames, LenH(Tmp_DCNames) - 1)
            End If
            
        itmX.SubItems(Col_Idx) = "" & Tmp_DCNames: Col_Idx = Col_Idx + 1    '���γ���
        
        
        
        
        
        
        
        
        
        TOTAL_FEE = TOTAL_FEE + Val(rs!TOTAL_PAID)
        
        
        TotalSum = TotalSum + Val("" & rs!TOTAL_PARKING_PAYMENT)
        DcSum = DcSum + Val("" & rs!DC_MONEY)
        If rs!Gubun = "����" Then
            SumCash = SumCash + Val(rs!TOTAL_PAID)
        ElseIf rs!Gubun = "�ſ�ī��" Then
            SumCard = SumCard + Val(rs!TOTAL_PAID)
        Else '���� + �ſ�ī��
            SumCash = SumCash + Val(rs!CASH_PAID)
            SumCard = SumCard + Val(rs!CARD_PAID)
        End If
        
        
        rs.MoveNext
        INDEX_NO = INDEX_NO + 1
    Loop
    Set rs = Nothing
    
    
    lbl_TotalSum.Caption = TotalSum
    lbl_DcSum.Caption = DcSum
    lbl_RealSum.Caption = TOTAL_FEE
    
    lbl_Search.Caption = TOTAL_FEE
    lbl_Cash.Caption = SumCash
    lbl_Card.Caption = SumCard
    
    lbl_Search.Caption = Format(TOTAL_FEE, "#,###,###,##0") & "��"

End Sub

Private Sub ListView_REG_DrawWebDC()
    Dim Column_to_size As Integer
    With Me
        Call ListViewExtended(.ListView_REG)
        .ListView_REG.View = lvwReport
        .ListView_REG.ListItems.Clear
        .ListView_REG.ColumnHeaders.Clear
        .ListView_REG.ColumnHeaders.Add , , " No  "
        .ListView_REG.ColumnHeaders.Add , , " �����ڵ�        "
        .ListView_REG.ColumnHeaders.Add , , " ��Ʈ��          "
        .ListView_REG.ColumnHeaders.Add , , " ������ȣ              "
        '.ListView_REG.ColumnHeaders.Add , , " ���α���             "
        .ListView_REG.ColumnHeaders.Add , , " ���γ���              "
        .ListView_REG.ColumnHeaders.Add , , " ����Ͻ�                        "
        .ListView_REG.ColumnHeaders.Add , , ""
        For Column_to_size = 0 To .ListView_REG.ColumnHeaders.Count - 2
             SendMessage .ListView_REG.hwnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
        Next
    End With
End Sub
Private Sub ListView_REG_SQLWebDC(ByVal sQry As String)
    Dim rs As Recordset
    Dim itmX As ListItem
    Dim iRow As Long
    Dim iCol As Long
    Dim iCount As Long
    Dim iTotal As Long
    Dim iFreeCount As Long
    
    
    iCount = 0
    iTotal = 0
    iRow = 1
    
    Set rs = New ADODB.Recordset
    rs.Open sQry, adoConn

    Do While Not (rs.EOF)
        Set itmX = ListView_REG.ListItems.Add(, , "" & iRow)
        
        iCol = 1
        itmX.SubItems(iCol) = "" & rs!PCODE: iCol = iCol + 1
        itmX.SubItems(iCol) = "" & rs!PNAME: iCol = iCol + 1
        itmX.SubItems(iCol) = "" & rs!DC_CARNO: iCol = iCol + 1
        itmX.SubItems(iCol) = "" & rs!DC_NAME: iCol = iCol + 1
        itmX.SubItems(iCol) = "" & rs!DT_DC: iCol = iCol + 1
        
        If (rs!DC_CODE = 99999) Then
            iFreeCount = iFreeCount + 1
        Else
            iTotal = iTotal + rs!DC_CODE
        End If
        iRow = iRow + 1
        
        rs.MoveNext
        
    Loop
    Set rs = Nothing
    lbl_COUNT = "" & CStr(iRow - 1) & " ��"
    
    'lbl_Search = "" + CStr(iTotal) & " ��"
    If (iTotal >= 60) Then
        lbl_Search = "" & Int(iTotal / 60) & "�ð�  "
        If (iTotal Mod 60) Then
            lbl_Search = lbl_Search & (iTotal Mod 60) & "��"
        End If
    Else
        lbl_Search = "" & iTotal & " �� "
    End If
    
    lbl_Search = "����:" & iFreeCount & "��, " & lbl_Search
End Sub
Private Sub ListView_Reg_ClearCaptionWebDC()
    lbl_title(0).Caption = "# ���� ����"
    lbl_Name(0).Caption = "��Ʈ�ʸ�:"
    lbl_Name(1).Caption = "������ȣ:"
    lbl_Name(3).Caption = "���γ���:"
    lbl_Name(11).Caption = "����Ͻ�:"
    lbl_Name(2).Caption = ""
    lbl_Name(4).Caption = ""
    
    lbl_Name(5).Caption = ""
    lbl_Name(6).Caption = ""
    lbl_Name(8).Caption = ""
    lbl_Name(10).Caption = ""
    lbl_Name(7).Caption = ""
    lbl_Name(9).Caption = ""
End Sub

Public Sub ListView_REG_Draw()
Dim Column_to_size As Integer

With Me
    Call ListViewExtended(.ListView_REG)
'''    .ListView_REG.View = lvwReport
'''    .ListView_REG.ListItems.Clear
'''    .ListView_REG.ColumnHeaders.Clear
'''    .ListView_REG.ColumnHeaders.Add , , " No  "
'''    .ListView_REG.ColumnHeaders.Add , , " ��������         "
'''    .ListView_REG.ColumnHeaders.Add , , " �����ð�     "
'''    .ListView_REG.ColumnHeaders.Add , , " ��������         "
'''    .ListView_REG.ColumnHeaders.Add , , " �����ð�     "
'''    .ListView_REG.ColumnHeaders.Add , , " ������ȣ / ���ڵ�     "
'''    .ListView_REG.ColumnHeaders.Add , , " �� �����ð�(��)  "
'''    .ListView_REG.ColumnHeaders.Add , , " �� ���νð�(��)  "
'''    '.ListView_REG.ColumnHeaders.Add , , " ��������(��)  "
'''    .ListView_REG.ColumnHeaders.Add , , " �������       "
'''    .ListView_REG.ColumnHeaders.Add , , " �� ��          "
'''    .ListView_REG.ColumnHeaders.Add , , " ��������       "
'''    .ListView_REG.ColumnHeaders.Add , , " �Ǽ��ɾ�       "
'''    '.ListView_REG.ColumnHeaders.Add , , " POS_CODE    "
'''    .ListView_REG.ColumnHeaders.Add , , " "

    .ListView_REG.View = lvwReport
    .ListView_REG.ListItems.Clear
    .ListView_REG.ColumnHeaders.Clear
    .ListView_REG.ColumnHeaders.Add , , " No  "
    .ListView_REG.ColumnHeaders.Add , , " Ƽ���ڵ�        "
    .ListView_REG.ColumnHeaders.Add , , " ��������      "
    .ListView_REG.ColumnHeaders.Add , , " �����ð�  "
    .ListView_REG.ColumnHeaders.Add , , " ��������      "
    .ListView_REG.ColumnHeaders.Add , , " �����ð�  "
    '.ListView_REG.ColumnHeaders.Add , , " ������ȣ / ���ڵ�     "
    .ListView_REG.ColumnHeaders.Add , , " ������ȣ               "
    .ListView_REG.ColumnHeaders.Add , , " �����ð�               "
    .ListView_REG.ColumnHeaders.Add , , " ���νð�  "
    .ListView_REG.ColumnHeaders.Add , , " ���αݾ�   "
    '.ListView_REG.ColumnHeaders.Add , , " ��������       "
    .ListView_REG.ColumnHeaders.Add , , " �������       "
    '.ListView_REG.ColumnHeaders.Add , , " ����/�ſ�ī��  "
    .ListView_REG.ColumnHeaders.Add , , " �� ��       "
    
    
    '.ListView_REG.ColumnHeaders.Add , , " ��������       "
    .ListView_REG.ColumnHeaders.Add , , " �Ǽ��ɾ�       "
    '.ListView_REG.ColumnHeaders.Add , , " POS_CODE    "
    .ListView_REG.ColumnHeaders.Add , , " ���γ���       "
    .ListView_REG.ColumnHeaders.Add , , " "
    
    
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
Dim Tmp_Path1, Tmp_Path2 As String

ListView_REG.SetFocus

If (sSelButton = "�� ��") Then
    lbl_Name(5) = ListView_REG.SelectedItem.SubItems(2) & " " & ListView_REG.SelectedItem.SubItems(3)
    lbl_Name(6) = ListView_REG.SelectedItem.SubItems(4) & " " & ListView_REG.SelectedItem.SubItems(5)
    'lbl_Name(7) = ListView_REG.SelectedItem.SubItems(10) & " ��"
    lbl_Name(7) = ListView_REG.SelectedItem.SubItems(10) & ""
    'lbl_Name(8) = ListView_REG.SelectedItem.SubItems(7) & " ��"
    lbl_Name(8) = ListView_REG.SelectedItem.SubItems(7) & ""
    lbl_Name(9) = ListView_REG.SelectedItem.SubItems(8) & " ��"
    lbl_Name(10) = ListView_REG.SelectedItem.SubItems(8) & " ��"
    
    Dim nRealMoney As Long
    nRealMoney = GetOnlyNumber(ListView_REG.SelectedItem.SubItems(12))
    If (ListView_REG.SelectedItem.SubItems(11) = "�ſ�ī��" And nRealMoney > 0) Then
        cmd_Button(4).Enabled = True '�ſ�ī�� ������ҹ�ư
    Else
        cmd_Button(4).Enabled = False
    End If
    
    
    lbl_seq = ListView_REG.SelectedItem '������ȣ
    lbl_ticketcode = ListView_REG.SelectedItem.SubItems(1) 'Ƽ���ڵ�
    lbl_carno = ListView_REG.SelectedItem.SubItems(6) '������ȣ
    
    cmd_Button(5).Enabled = True '����������� ��ư
    
Else '������ �˻�
    lbl_Name(5).Caption = ListView_REG.SelectedItem.SubItems(2)
    lbl_Name(6).Caption = ListView_REG.SelectedItem.SubItems(3)
    lbl_Name(8).Caption = ListView_REG.SelectedItem.SubItems(4)
    lbl_Name(10).Caption = ListView_REG.SelectedItem.SubItems(5)
    
    cmd_Button(5).Enabled = False '����������� ��ư
End If

End Sub

Function GetOnlyNumber(ByVal sString As String)
    Dim i As Integer
    Dim Num As String

    For i = Len(sString) To 1 Step -1
        If IsNumeric(Mid(sString, i, 1)) Then
            Num = Mid(sString, i, 1) & Num
        ElseIf Mid(sString, i, 1) = "-" Then
            Num = Mid(sString, i, 1) & Num
        End If
    Next i
    GetOnlyNumber = Num

End Function
    
    
Public Sub ListView1_SQL(ByVal sQry As String)
Dim rs As ADODB.Recordset
Dim itmX As ListItem
Dim INDEX_NO As Long
Dim TOTAL_FEE As Long
Dim SumCash As Long
Dim SumCard As Long
Dim DcSum As Long
Dim TotalSum As Long

INDEX_NO = 1
TOTAL_FEE = 0
Set rs = New ADODB.Recordset
'rs.Open CardQry, adoConn
rs.Open sQry, adoConn
lbl_COUNT = rs.RecordCount
Do While Not (rs.EOF)
    Set itmX = ListView_REG.ListItems.Add(, , "" & INDEX_NO)
    itmX.SubItems(1) = "" & rs!TICKET_CODE
    itmX.SubItems(2) = "" & rs!TrdDate
    itmX.SubItems(3) = "" & rs!CardKind
    itmX.SubItems(4) = "" & rs!OrgNm
    itmX.SubItems(5) = "" & rs!TrdMoney
    itmX.SubItems(6) = "" & rs!carnum
    itmX.SubItems(7) = "" & rs!REG_DATE
    'TotalSum = TotalSum + Val(rs!TrdMoney)
    TOTAL_FEE = TOTAL_FEE + Val(rs!TrdMoney)
    
'    lbl_TotalSum.Caption = TotalSum
'    lbl_DcSum.Caption = DcSum
'    lbl_RealSum.Caption = TOTAL_FEE
    
    lbl_Search.Caption = TOTAL_FEE
'    lbl_Cash.Caption = SumCash
'    lbl_Card.Caption = SumCard
    
    rs.MoveNext
    INDEX_NO = INDEX_NO + 1
Loop
Set rs = Nothing

End Sub

Public Sub ListView1_Draw()
Dim Column_to_size As Integer

With Me
    'Call ListViewExtended(.ListView1)
    Call ListViewExtended(.ListView_REG)
    .ListView_REG.View = lvwReport
    .ListView_REG.ListItems.Clear
    .ListView_REG.ColumnHeaders.Clear
    .ListView_REG.ColumnHeaders.Add , , " No  "
    .ListView_REG.ColumnHeaders.Add , , " TicketCode                 "
    .ListView_REG.ColumnHeaders.Add , , " �����Ͻ�           "
    .ListView_REG.ColumnHeaders.Add , , " ī������                "
    .ListView_REG.ColumnHeaders.Add , , " ī���           "
    .ListView_REG.ColumnHeaders.Add , , " ����ݾ�          "
    .ListView_REG.ColumnHeaders.Add , , " ������ȣ          "
    .ListView_REG.ColumnHeaders.Add , , " RegDate                   "
    '.ListView1.ColumnHeaders.Add , , " "
    For Column_to_size = 0 To .ListView_REG.ColumnHeaders.Count - 2
         SendMessage .ListView_REG.hwnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
    Next
End With
End Sub

Private Sub cmd_Button_Click(Index As Integer)
    Dim myExcelFile As New ExcelFile
    Dim tmpFileName As String
    Dim sql_str As String
    Dim sSDT As String
    Dim sEDT As String

    cmd_Button(4).Enabled = False
    
Select Case Index
    Case 0  '����
        List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    �������� ����", 0
        Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    �������� ����")
        Unload Me
        'Me.Hide
        Exit Sub

    Case 1

        tmpFileName = Format(Now, "YYYYMMDD_HHMMSS")
        tmpFileName = App.Path & "\Excel\" & tmpFileName & "_��������" & ".xls"
        'Call makeexcel(ListView_REG, tmpFileName, "_��������")
        Call MakeCSV(ListView_REG, tmpFileName)
        Exit Sub
        
    Case 2
        '������������ �˻�
        sSelButton = "�� ��"
        Me.MousePointer = 11
        Glo_SQL_SEARCH = ""

        cmd_Button(5).Enabled = False '����������� ��ư
        
        '���� ����
        '��ȸ�Ⱓ
        If (cmb_Search.text = "��ü" Or cmb_Search.text = "�ſ�ī��") Then
        
        
            
            
            sql_str = "SELECT * FROM tb_calculate WHERE ( REG_DATE >= '" & Format(DTPicker1, "yyyy-mm-dd") & " " & Format(DTPicker3, "hh:nn:ss") & "') AND ( REG_DATE <= '" & Format(DTPicker2, "yyyy-mm-dd") & " " & Format(DTPicker4, "hh:nn:ss") & "')"
            '������ȣ �˻�
            If (txt_CarNo.text <> "") Then
                If IsNumeric(txt_CarNo) And Len(txt_CarNo) = 4 Then
                Else
                    MsgBox "������ȣ ��4�ڸ��� Ȯ�����ּ���."
                    txt_CarNo.text = ""
                    Me.MousePointer = 0
                    Exit Sub
                End If
                sql_str = sql_str & " AND (TICKET_NO Like '%" & txt_CarNo.text & "')"
            End If
            
            If (cmb_Search.text <> "��ü") Then
                sql_str = sql_str & " AND (GUBUN Like '" & cmb_Search & "')"
            End If
'            If (Len(cmb_DCGubun.text) <> 0) Then
'                sql_str = sql_str & " AND (DC_GUBUN Like '" & cmb_DCGubun & "')"
'            End If
            If (cmb_DCGubun.text <> "��ü") Then
                If (cmb_DCGubun.text = "������") Then
                    sql_str = sql_str & " AND (DC_GUBUN Like '" & 0 & "')"
                Else
                    sql_str = sql_str & " AND (DC_GUBUN Like '" & cmb_DCGubun & "')"
                End If
            End If
            
            
            sql_str = sql_str & " ORDER BY REG_DATE"
            Glo_SQL_REG = sql_str
            Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & sql_str)
            
            Call ListView_REG_Draw
            Call ListView_REG_SQL
            Call ListView_Reg_ClearCaption
            Me.MousePointer = 0
'        ElseIf (cmb_Search.text = "�ſ�ī��") Then
'            sql_str = "SELECT * FROM tb_kicc_log WHERE ( Reg_Date >= '" & Format(DTPicker1, "yyyy-mm-dd") & " 00:00:00') AND ( Reg_Date <= '" & Format(DTPicker2, "yyyy-mm-dd") & " 23:59:59')"
'            '������ȣ �˻�
'            If (txt_CarNo.text <> "") Then
'                If IsNumeric(txt_CarNo) And Len(txt_CarNo) = 4 Then
'                Else
'                    MsgBox "������ȣ ��4�ڸ��� Ȯ�����ּ���."
'                    Me.MousePointer = 0
'                    Exit Sub
'                End If
'                sql_str = sql_str & " AND (CarNum Like '%" & txt_CarNo.text & "')"
'            End If
'
'            sql_str = sql_str & " ORDER BY Reg_Date"
'            'CardQry = sql_str
'            Call DataLogger(sql_str)
'
'            Call ListView1_Draw
'            Call ListView1_SQL(sql_str)
'            Me.MousePointer = 0
        End If
        Exit Sub
        
    Case 3
        sSelButton = "������ �˻�"
        Me.MousePointer = 11
        
        sSDT = Format(DTPicker1, "yyyy-mm-dd") & " " & Format(DTPicker3, "hh:nn:ss")
        sEDT = Format(DTPicker2, "yyyy-mm-dd") & " " & Format(DTPicker4, "hh:nn:ss")
        sql_str = "SELECT * FROM tb_web_dc WHERE ( DT_DC >= '" & sSDT & "') AND ( DT_DC <= '" & sEDT & "') "
        
        If (Len(cmb_Partner.text) > 0) Then
            If (cmb_Partner.text = "��ü") Then
            Else
                sql_str = sql_str + " AND (PNAME = '" & cmb_Partner.text & "' ) "
            End If
        End If
        
        If (txt_CarNo.text <> "") Then
            If IsNumeric(txt_CarNo) And Len(txt_CarNo) = 4 Then
            Else
                MsgBox "������ȣ ��4�ڸ��� Ȯ�����ּ���."
                txt_CarNo.text = ""
                Me.MousePointer = 0
                Exit Sub
            End If
            sql_str = sql_str & " AND (DC_CARNO Like '%" & txt_CarNo.text & "')"
        End If
        
        sql_str = sql_str + " ORDER BY SEQ"
        
        Call ListView_REG_DrawWebDC
        Call ListView_REG_SQLWebDC(sql_str)
        Call ListView_Reg_ClearCaptionWebDC
        Me.MousePointer = 0
        Exit Sub
    
    
    Case 4  'ī��������
        MBox.Label3.FontSize = 20
        MBox.Label3.Caption = lbl_carno
        MBox.Label2.Caption = "ī��������"
        MBox.Label1.Caption = "�� ������ ī�� ������ ����մϴ�." & vbCrLf & "��� �����Ͻðڽ��ϱ�?"
        MBox.Show 1
        If (Glo_MsgRet = True) Then
            Call DataLogger("[ī��������] " & "�ε���:" & lbl_seq & ", Ƽ���ڵ�:" & lbl_ticketcode & "," & lbl_carno & " ī�������� �մϴ�")
            Glo_APSCMD_Str = CM_CARDCANCEL & lbl_ticketcode
            'Glo_APSCMD_Str = CM_CARDCANCEL & lbl_seq
            
            If (CardCancel_Sock.State <> sckClosed) Then
                CardCancel_Sock.Close
            End If
            CardCancel_Sock.Connect Glo_Aps_IP, Glo_Aps_PORT '����ȭ��(5889)���� ó����
            
        End If
        
        Exit Sub
    
    
    Case 5
        Call DataLogger("[�����������] " & "Ƽ���ڵ�:" & lbl_ticketcode & "," & lbl_carno & " ����� �մϴ�")
        Glo_APSCMD_Str = CM_REPRINT & lbl_ticketcode
        
        If (RePrint_Sock.State <> sckClosed) Then
            RePrint_Sock.Close
        End If
        RePrint_Sock.Connect Glo_Aps_IP, Glo_Aps_PORT '����ȭ��(5889)���� ó����
            
        Exit Sub

End Select
    Call DataLogger("[��������] " & "����:��ư(" & Index & ")")
    On Error Resume Next
End Sub

'����Ű �Է½� �� ����
'���Ӽ� keypreview = true ����
Private Sub Form_KeyPress(KeyAscii As Integer)
    
If KeyAscii = vbKeyReturn Then
    Call cmd_Button_Click(2)
    KeyAscii = 0
    'SendKeys "{TAB}"
End If

End Sub


Private Sub RePrint_Sock_Connect()
    Dim sdata As String
    Dim bData() As Byte
    Dim i As Integer
    
On Error GoTo Err_P
    
    If (Len(Glo_APSCMD_Str) > 0) Then
        sdata = Glo_APSCMD_Str
        ReDim bData(Len(sdata) - 1) As Byte
        
        bData = StrConv(sdata, vbFromUnicode)
        RePrint_Sock.SendData bData
        
        Call DataLogger("[�����������]  SND : " & Glo_APSCMD_Str)
        Glo_APSCMD_Str = ""
    End If
    
    Exit Sub

Err_P:
    Call DataLogger(" [�����������] Connect Err_Msg : " & Err.Description)
End Sub

Private Sub RePrint_Sock_DataArrival(ByVal bytesTotal As Long)
    Dim rMsg As String
    Dim B() As Byte
    Dim Ret As Integer
    Dim i As Integer
    Dim sdata As String
    
    On Error GoTo Err_P
    
    ReDim B(bytesTotal - 1)
    
    RePrint_Sock.GetData B(), vbArray + vbByte, bytesTotal
    For i = 0 To bytesTotal - 1
        If (B(i) >= &H80) Then
            rMsg = rMsg & Chr$(Val("&H" & Hex(B(i)) & Hex(B(i + 1))))
            i = i + 1
        Else
            rMsg = rMsg & Chr$(B(i))
        End If
    Next i
    
    Call DataLogger("[�����������]  " & lbl_carno & ":" & rMsg)
    'List1.AddItem Format(Now, "HH:NN:SS") & " RCV : " & rMsg, 0
    If (InStr(rMsg, "CMD_SUCCESS") > 0) Then
        List1.AddItem Format(Now, "HH:NN:SS") & " [�����������] ����Ϸ� : " & lbl_carno, 0
    End If
    
    RePrint_Sock.Close
    
    Exit Sub
    
Err_P:
        Call DataLogger("[�����������] Recv Err_Msg : " & Err.Description)
End Sub


Private Sub CardCancel_Sock_Connect()
    Dim sdata As String
    Dim bData() As Byte
    Dim i As Integer
    
On Error GoTo Err_P
    
    If (Len(Glo_APSCMD_Str) > 0) Then
        sdata = Glo_APSCMD_Str
        ReDim bData(Len(sdata) - 1) As Byte
        
        bData = StrConv(sdata, vbFromUnicode)
        CardCancel_Sock.SendData bData
        
        Call DataLogger("[ī��������]  SND : " & Glo_APSCMD_Str)
        Glo_APSCMD_Str = ""
    End If
    
    Exit Sub

Err_P:
    Call DataLogger(" [ī��������] Connect Err_Msg : " & Err.Description)
    Call DataLogger(" [ī��������] ��Ʈ��ũ ����!")
End Sub

Private Sub CardCancel_Sock_DataArrival(ByVal bytesTotal As Long)
    Dim rMsg As String
    Dim B() As Byte
    Dim Ret As Integer
    Dim i As Integer
    Dim sdata As String
    
    On Error GoTo Err_P
    
    ReDim B(bytesTotal - 1)
    
    CardCancel_Sock.GetData B(), vbArray + vbByte, bytesTotal
    For i = 0 To bytesTotal - 1
        If (B(i) >= &H80) Then
            rMsg = rMsg & Chr$(Val("&H" & Hex(B(i)) & Hex(B(i + 1))))
            i = i + 1
        Else
            rMsg = rMsg & Chr$(B(i))
        End If
    Next i
    
    Call DataLogger("[ī��������]  " & lbl_carno & ":" & rMsg)
    'List1.AddItem Format(Now, "HH:NN:SS") & " RCV : " & rMsg, 0
    If (InStr(rMsg, "CMD_SUCCESS") > 0) Then
        List1.AddItem Format(Now, "HH:NN:SS") & " [ī��������] ��ҿϷ� : " & lbl_carno, 0
        Call DataLogger("[ī��������] ��ҿϷ� : " & lbl_carno)
    End If
    
    RePrint_Sock.Close
    
    Exit Sub
    
Err_P:
        Call DataLogger("[ī��������] Recv Err_Msg : " & Err.Description)
End Sub


