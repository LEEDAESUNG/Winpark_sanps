VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmTicketIn 
   BorderStyle     =   1  '���� ����
   Caption         =   " �Ϲݱ� ���� ��Ȳ"
   ClientHeight    =   14715
   ClientLeft      =   300
   ClientTop       =   1800
   ClientWidth     =   19200
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmTicketIn.frx":0000
   ScaleHeight     =   14715
   ScaleWidth      =   19200
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
      Left            =   12930
      TabIndex        =   28
      Top             =   2130
      Width           =   1335
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "�������"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   915
      Left            =   120
      TabIndex        =   0
      Top             =   13680
      Width           =   18975
   End
   Begin ComctlLib.ListView ListView_REG 
      Height          =   5325
      Left            =   360
      TabIndex        =   1
      Top             =   2910
      Width           =   18450
      _ExtentX        =   32544
      _ExtentY        =   9393
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
      TabIndex        =   2
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
      Picture         =   "FrmTicketIn.frx":F4A3
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   660
      Index           =   1
      Left            =   15450
      TabIndex        =   3
      Top             =   1950
      Width           =   1410
      _Version        =   65536
      _ExtentX        =   2487
      _ExtentY        =   1164
      _StockProps     =   78
      Caption         =   "�� ��"
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
      Picture         =   "FrmTicketIn.frx":F7F4
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   660
      Index           =   2
      Left            =   16920
      TabIndex        =   4
      Top             =   1950
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
      Picture         =   "FrmTicketIn.frx":FB45
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   345
      Left            =   6030
      TabIndex        =   5
      Top             =   2130
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
      Format          =   16711680
      CurrentDate     =   36927
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   345
      Left            =   8820
      TabIndex        =   6
      Top             =   2130
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
      Format          =   16711680
      CurrentDate     =   36927
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   660
      Index           =   3
      Left            =   16080
      TabIndex        =   29
      Top             =   690
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
      Picture         =   "FrmTicketIn.frx":FE96
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
      Left            =   11790
      TabIndex        =   27
      Top             =   2175
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
      Index           =   9
      Left            =   10950
      TabIndex        =   26
      Top             =   2175
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
      Left            =   19890
      TabIndex        =   25
      Top             =   8580
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
      Index           =   7
      Left            =   8160
      TabIndex        =   24
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
      Left            =   4920
      TabIndex        =   23
      Top             =   2175
      Width           =   1035
   End
   Begin VB.Label lbl_title 
      BackStyle       =   0  '����
      Caption         =   "�Ϲݱ� ���� ��Ȳ �˻�"
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
      TabIndex        =   22
      Top             =   2010
      Width           =   4065
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '����
      Caption         =   "��ȸ �Ǽ� :"
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
      Left            =   15840
      TabIndex        =   21
      Top             =   8550
      Width           =   1215
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
      Left            =   17100
      TabIndex        =   20
      Top             =   8550
      Width           =   1425
   End
   Begin VB.Label lbl_title 
      BackStyle       =   0  '����
      Caption         =   "# ���� ���� ����"
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
      TabIndex        =   19
      Top             =   9480
      Width           =   4035
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '����
      Caption         =   "�Ϲݱ� ���� ��Ȳ"
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
      TabIndex        =   18
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
      Left            =   21000
      TabIndex        =   17
      Top             =   8550
      Width           =   2025
   End
   Begin VB.Image ImageIn 
      Appearance      =   0  '���
      BorderStyle     =   1  '���� ����
      Height          =   4290
      Index           =   0
      Left            =   240
      Picture         =   "FrmTicketIn.frx":101E7
      Stretch         =   -1  'True
      Top             =   9300
      Width           =   5730
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
      TabIndex        =   16
      Top             =   10500
      Width           =   1665
   End
   Begin VB.Label lbl_Name 
      BackStyle       =   0  '����
      Caption         =   "�νĹ�ȣ :"
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
      TabIndex        =   15
      Top             =   11025
      Width           =   1665
   End
   Begin VB.Label lbl_Name 
      BackStyle       =   0  '����
      Caption         =   "�νĻ��� :"
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
      TabIndex        =   14
      Top             =   12075
      Width           =   1665
   End
   Begin VB.Label lbl_Name 
      BackStyle       =   0  '����
      Caption         =   "������ȣ :"
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
      TabIndex        =   13
      Top             =   11550
      Width           =   1665
   End
   Begin VB.Label lbl_Name 
      BackStyle       =   0  '����
      Caption         =   "ó���Ͻ� :"
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
      TabIndex        =   12
      Top             =   12600
      Width           =   1665
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
      Left            =   14070
      TabIndex        =   11
      Top             =   10500
      Width           =   4665
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
      Left            =   14070
      TabIndex        =   10
      Top             =   11025
      Width           =   4665
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
      Left            =   14070
      TabIndex        =   9
      Top             =   12075
      Width           =   4665
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
      Left            =   14070
      TabIndex        =   8
      Top             =   11550
      Width           =   4665
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
      ForeColor       =   &H00C0FFFF&
      Height          =   435
      Index           =   9
      Left            =   14070
      TabIndex        =   7
      Top             =   12600
      Width           =   4665
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H006F3C2F&
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00C0C0C0&
      Height          =   975
      Left            =   240
      Top             =   1770
      Width           =   18735
   End
End
Attribute VB_Name = "FrmTicketIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim i As Integer

Left = (Screen.Width - Width) / 2   ' ���� ���η� �߾ӿ� �����ϴ�.
Top = (Screen.Height - Height) / 2   ' ���� ���η� �߾ӿ� �����ϴ�.
'Left = 0
'Top = 0
'Me.cmb_Gubun = Me.cmb_Gubun.List(0)
DTPicker1.value = Now
DTPicker2.value = Now
Glo_SQL_REG = "SELECT * FROM ilbancarin WHERE (ó���Ͻ� >= '" & Format(DTPicker1, "yyyymmdd") & "000000') AND (ó���Ͻ� <= '" & Format(DTPicker2, "yyyymmdd") & "235959') ORDER BY ó���Ͻ�"
'Glo_SQL_REG = "SELECT * From TB_FEE WHERE ORDER BY REG_DATE ASC"
Call ListView_REG_Draw
Call ListView_REG_SQL
List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    �Ϲݱ� ������Ȳ ����...!!", 0
Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    �Ϲݱ� ������Ȳ ����...!!")


End Sub

Public Sub ListView_REG_SQL()
Dim rs As Recordset
Dim QRY As String
Dim itmX As ListItem
Dim INDEX_NO As Long
Dim TOTAL_FEE As Single

INDEX_NO = 1
TOTAL_FEE = 0
Set rs = New ADODB.Recordset
rs.Open Glo_SQL_REG, adoConn
lbl_COUNT = rs.RecordCount
Do While Not (rs.EOF)
    Set itmX = ListView_REG.ListItems.Add(, , "" & INDEX_NO)
    itmX.SubItems(1) = "" & rs!��������
    itmX.SubItems(2) = "" & rs!�����ð�
    itmX.SubItems(3) = "" & rs!�νĹ�ȣ
    itmX.SubItems(4) = "" & rs!������ȣ
    itmX.SubItems(5) = "" & rs!�νĻ���
    itmX.SubItems(6) = "" & rs!ó���Ͻ�
    itmX.SubItems(7) = "" & rs!����
    itmX.SubItems(8) = "" & rs!����Ʈ����
    itmX.SubItems(9) = "" & rs!�̹�����
'    TOTAL_FEE = TOTAL_FEE + Val(rs!�Ǽ��ɱݾ�)
'    lbl_Search.Caption = TOTAL_FEE & "��"
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
    .ListView_REG.ColumnHeaders.Add , , " ��������         "
    .ListView_REG.ColumnHeaders.Add , , " �����ð�     "
    .ListView_REG.ColumnHeaders.Add , , " �νĹ�ȣ         "
    .ListView_REG.ColumnHeaders.Add , , " ������ȣ         "
    .ListView_REG.ColumnHeaders.Add , , " �νĻ���     "
    .ListView_REG.ColumnHeaders.Add , , " ó���Ͻ�                           "
    .ListView_REG.ColumnHeaders.Add , , " ��������         "
    .ListView_REG.ColumnHeaders.Add , , " ��Ϻμ�         "
    .ListView_REG.ColumnHeaders.Add , , " �̹�����                           "
    .ListView_REG.ColumnHeaders.Add , , " "
    
    For Column_to_size = 0 To .ListView_REG.ColumnHeaders.Count - 2
         SendMessage .ListView_REG.hWnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
    Next
End With
End Sub

Private Sub ListView_REG_ItemClick(ByVal Item As ComctlLib.ListItem)
Dim Tmp_Path1, Tmp_Path2 As String

ListView_REG.SetFocus

lbl_Name(5) = ListView_REG.SelectedItem.SubItems(1) & " " & ListView_REG.SelectedItem.SubItems(2)
lbl_Name(6) = ListView_REG.SelectedItem.SubItems(3)
lbl_Name(7) = ListView_REG.SelectedItem.SubItems(5)
lbl_Name(8) = ListView_REG.SelectedItem.SubItems(4)
lbl_Name(9) = ListView_REG.SelectedItem.SubItems(6)

Tmp_Path1 = Dir(ListView_REG.SelectedItem.SubItems(9))
If (Tmp_Path1 = "") Then
    ImageIn(0).Picture = Nothing
Else
    ImageIn(0).Picture = LoadPicture(ListView_REG.SelectedItem.SubItems(9))
End If
'Tmp_Path2 = Dir(ListView_REG.SelectedItem.SubItems(9))
'If (Tmp_Path2 = "") Then
'    ImageIn(1).Picture = Nothing
'Else
'    ImageIn(1).Picture = LoadPicture(ListView_REG.SelectedItem.SubItems(9))
'End If

End Sub

Private Sub cmd_Button_Click(Index As Integer)
Dim myExcelFile As New ExcelFile
Dim tmpFileName As String
Dim sql_str As String

On Error Resume Next

Select Case Index
    Case 0  '����
        List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    �Ϲݱ� ������Ȳ ����", 0
        Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    �Ϲݱ� ������Ȳ ����")
        Unload Me
        Exit Sub

    Case 1  '����
        If (Len(lbl_Name(9)) = 0) Then
            Msg_Box.Label2.Caption = "������ ���� ����"
            Msg_Box.Label1.Caption = "������ �����͸� ������ �ֽʽÿ�."
            Msg_Box.Show 1
            Exit Sub
        End If
        
        MBox.Label3.Caption = lbl_Name(8)
        MBox.Label1.Caption = "�� �Ϲ������� �����ڷḦ �����մϴ�. �����Ͻðڽ��ϱ�?"
        MBox.Label2.Caption = "�Ϲݱ� �ڷ� ����"
        MBox.Show 1
        
        If (Glo_MsgRet = True) Then
            '��������
            adoConn.Execute "Delete from ilbancarin where ó���Ͻ� = '" & lbl_Name(9) & "'"
            List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & lbl_Name(5) & "  " & lbl_Name(8) & " �Ϲ������� ���������� ����", 0
            Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & lbl_Name(5) & "  " & lbl_Name(8) & " �Ϲ������� ���������� ����")
        Else
        End If
        Call cmd_Button_Click(2)
        Exit Sub
        
    Case 2
        '������������ �˻�
        Me.MousePointer = 11
        Glo_SQL_SEARCH = ""
        '���� ����
        '��ȸ�Ⱓ
        'Glo_SQL_REG = "SELECT * FROM TB_FEE WHERE (REG_DATE >= '" & Format(DTPicker1, "yyyy-mm-dd") & " 00:00:00') AND (REG_DATE <= '" & Format(DTPicker2, "yyyy-mm-dd") & " 23:59:59') ORDER BY REG_DATE ASC"
        sql_str = "SELECT * FROM ilbancarin WHERE (ó���Ͻ� >= '" & Format(DTPicker1, "yyyymmdd") & "000000" & "') AND (ó���Ͻ� <= '" & Format(DTPicker2, "yyyymmdd") & "235959" & "')"
        '������ȣ �˻�
        If (txt_CarNo.Text <> "") Then
            If IsNumeric(txt_CarNo) And Len(txt_CarNo) = 4 Then
            Else
                MsgBox "������ȣ ��4�ڸ��� Ȯ�����ּ���."
                Me.MousePointer = 0
                Exit Sub
            End If
            sql_str = sql_str & " AND (������ȣ Like '%" & txt_CarNo.Text & "')"
        End If
'        '������ �̸� �˻�
'        If (txt_Name.Text <> "") Then
'            sql_str = sql_str & " AND (DRIVER_NAME Like '%" & txt_Name.Text & "%')"
'        End If
        sql_str = sql_str & " ORDER BY ó���Ͻ�"
        Glo_SQL_REG = sql_str
        'List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & sql_str, 0
        Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & sql_str)
        
        'Debug.Print Glo_SQL_REG
        Call ListView_REG_Draw
        Call ListView_REG_SQL
        Me.MousePointer = 0
        Exit Sub
    Case 3
        tmpFileName = Format(Now, "YYYYMMDD_HHMMSS")
        tmpFileName = App.Path & "\Excel\" & tmpFileName & "_�Ϲݱ�_��������" & ".xls"
        Call makeexcel(ListView_REG, tmpFileName, "�Ϲݱ�_��������")
        Exit Sub
End Select


End Sub

'����Ű �Է½� �� ����
'���Ӽ� keypreview = true ����
Private Sub Form_KeyPRESS(KeyAscii As Integer)
    
If KeyAscii = vbKeyReturn Then
    Call cmd_Button_Click(2)
    KeyAscii = 0
    'SendKeys "{TAB}"
End If

End Sub

