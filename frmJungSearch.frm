VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmJungSearch 
   BorderStyle     =   1  '���� ����
   Caption         =   "����� �߱޴���"
   ClientHeight    =   14655
   ClientLeft      =   27855
   ClientTop       =   765
   ClientWidth     =   19125
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmJungSearch.frx":0000
   ScaleHeight     =   14655
   ScaleWidth      =   19125
   Begin ComctlLib.ListView ListView1 
      Height          =   10005
      Left            =   510
      TabIndex        =   0
      Top             =   1740
      Width           =   18195
      _ExtentX        =   32094
      _ExtentY        =   17648
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
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
   Begin Threed.SSCommand Command1 
      Height          =   570
      Left            =   8370
      TabIndex        =   1
      Top             =   13380
      Width           =   1755
      _Version        =   65536
      _ExtentX        =   3096
      _ExtentY        =   1005
      _StockProps     =   78
      Caption         =   "�� ��"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
      Picture         =   "frmJungSearch.frx":F651
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   345
      Left            =   1950
      TabIndex        =   2
      Top             =   13470
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   12648447
      CalendarForeColor=   12582912
      CalendarTitleBackColor=   8421504
      CalendarTitleForeColor=   12632256
      CalendarTrailingForeColor=   8421504
      Format          =   16580608
      CurrentDate     =   36927
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   345
      Left            =   5055
      TabIndex        =   3
      Top             =   13470
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   12648447
      CalendarForeColor=   12582912
      CalendarTitleBackColor=   8421504
      CalendarTitleForeColor=   12632256
      CalendarTrailingForeColor=   8421504
      Format          =   16580608
      CurrentDate     =   36927
   End
   Begin Threed.SSCommand SSCommand2 
      Cancel          =   -1  'True
      Height          =   660
      Left            =   16920
      TabIndex        =   4
      Top             =   13380
      Width           =   1500
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1164
      _StockProps     =   78
      Caption         =   "�� ��(&X)"
      ForeColor       =   16777215
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
      Picture         =   "frmJungSearch.frx":F9A2
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   660
      Left            =   15210
      TabIndex        =   5
      Top             =   13380
      Width           =   1500
      _Version        =   65536
      _ExtentX        =   2646
      _ExtentY        =   1164
      _StockProps     =   78
      Caption         =   "��������"
      ForeColor       =   16777215
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
      Picture         =   "frmJungSearch.frx":FCF3
   End
   Begin VB.Label LblRecordCount 
      Alignment       =   1  '������ ����
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   345
      Left            =   10440
      TabIndex        =   8
      Top             =   960
      Width           =   1785
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '����
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   435
      Index           =   1
      Left            =   7290
      TabIndex        =   7
      Top             =   13530
      Width           =   705
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '����
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   435
      Index           =   0
      Left            =   4200
      TabIndex        =   6
      Top             =   13530
      Width           =   705
   End
End
Attribute VB_Name = "frmJungSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim excel_sql_str As String


Private Sub Form_Load()
Dim Record_Source As String
Dim i As Integer

'On Error GoTo err_P

'Left = (Screen.Width - Width) / 2   ' ���� ���η� �߾ӿ� �����ϴ�.
'Top = (Screen.Height - Height) / 2   ' ���� ���η� �߾ӿ� �����ϴ�.
Left = 0
Top = 0

DTPicker1.value = Now
DTPicker2.value = Now

'���ó�¥ �����͸�
Glo_JungSearch = "SELECT * FROM regcar WHERE (�߱޽ð� >= '" & Format(DTPicker1, "yyyy-mm-dd") & " 00:00:00') AND (�߱޽ð� <= '" & Format(DTPicker2, "yyyy-mm-dd") & " 23:59:59') ORDER BY �߱޽ð�"
'Debug.Print Glo_JungSearch

Call ListView_Draw

Exit Sub

err_P:
        MsgBox "������ ���̽� �������" & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & "��Ʈ�� �����ڿ��� ���� �ٶ��ϴ�." & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & "������ ���̽� ���������� �ڷ�˻� ����� �����Ҽ� �����ϴ�.", vbCritical

End Sub


Public Sub ListView_Draw()
Dim Column_to_size As Integer
Dim rs As Recordset
Dim QRY As String
Dim itmX As ListItem
Dim INDEX_NO As Long
    
    Call ListViewExtended(ListView1)
    ListView1.View = lvwReport
    ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , " No   "
    ListView1.ColumnHeaders.Add , , " ������ȣ           "
    ListView1.ColumnHeaders.Add , , " �� ��            "
    ListView1.ColumnHeaders.Add , , " �� ��            "
    ListView1.ColumnHeaders.Add , , " �� ��       "
    ListView1.ColumnHeaders.Add , , " ��ȭ��ȣ             "
    ListView1.ColumnHeaders.Add , , " �������  "
    ListView1.ColumnHeaders.Add , , " �� �� ��             "
    ListView1.ColumnHeaders.Add , , " �� �� ��             "
    ListView1.ColumnHeaders.Add , , " �� �� ��             "
    ListView1.ColumnHeaders.Add , , " ����Ͻ�             "
    
    
    For Column_to_size = 0 To ListView1.ColumnHeaders.Count - 1
         SendMessage ListView1.hWnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
    Next
 
    Set rs = New ADODB.Recordset
    rs.Open Glo_JungSearch, adoConn
    LblRecordCount = rs.RecordCount & " ��"

    INDEX_NO = 1

    Do While Not (rs.EOF)
        Set itmX = ListView1.ListItems.Add(, , "" & INDEX_NO)
        itmX.SubItems(1) = "" & rs!������ȣ
        itmX.SubItems(2) = "" & rs!����
        itmX.SubItems(3) = "" & rs!����
        itmX.SubItems(4) = "" & rs!�̸�
        itmX.SubItems(5) = "" & rs!��ȭ��ȣ
        itmX.SubItems(6) = "" & rs!�������
        itmX.SubItems(7) = "" & rs!�߱���
        itmX.SubItems(8) = "" & rs!������
        itmX.SubItems(9) = "" & rs!������
        itmX.SubItems(10) = "" & rs!�߱޽ð�
        rs.MoveNext
        INDEX_NO = INDEX_NO + 1
    Loop
    INDEX_NO = 0
    Set rs = Nothing

End Sub

Private Sub SSCommand1_Click()
Dim i, j As Integer
Dim myExcelFile As New ExcelFile
Dim tmpFileName As String
    
tmpFileName = Format(Now, "YYYYMMDD_HHMMSS")
tmpFileName = App.Path & "\Excel\" & tmpFileName & "_�߱޴��� �˻�����" & ".xls"
'Call makeexcel(ListView1, tmpFileName, "������������Ȳ")
Call makeexcel(ListView1, tmpFileName, "�߱޴��� �˻�����")

Exit Sub

End Sub

'����
Private Sub SSCommand2_Click()
Unload Me
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

Me.MousePointer = 11

Glo_JungSearch = ""

'���� ����
sql_str = "SELECT * FROM regcar WHERE (�߱޽ð� >= '" & Format(DTPicker1, "yyyy-mm-dd") & " 00:00:00') AND (�߱޽ð� <= '" & Format(DTPicker2, "yyyy-mm-dd") & " 23:59:59') ORDER BY �߱޽ð�"

'Debug.Print sql_str

Glo_JungSearch = sql_str

Call ListView_Draw

Me.MousePointer = 0

'On Error Resume Next

End Sub
