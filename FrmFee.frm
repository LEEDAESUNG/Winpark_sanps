VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmFee 
   BorderStyle     =   1  '���� ����
   Caption         =   " ������� �������� ��ȸ"
   ClientHeight    =   14715
   ClientLeft      =   4140
   ClientTop       =   1980
   ClientWidth     =   19200
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmFee.frx":0000
   ScaleHeight     =   14715
   ScaleWidth      =   19200
   Begin VB.TextBox txt_Name 
      BackColor       =   &H00FFFFFF&
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
      Left            =   4860
      TabIndex        =   1
      Top             =   12390
      Width           =   2115
   End
   Begin VB.TextBox txt_CarNo 
      BackColor       =   &H00FFFFFF&
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
      Left            =   4860
      TabIndex        =   0
      Top             =   11790
      Width           =   2115
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
      TabIndex        =   5
      Top             =   13350
      Width           =   18975
   End
   Begin ComctlLib.ListView ListView_REG 
      Height          =   8055
      Left            =   360
      TabIndex        =   7
      Top             =   2370
      Width           =   18450
      _ExtentX        =   32544
      _ExtentY        =   14208
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
      TabIndex        =   8
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
      Picture         =   "FrmFee.frx":F4A3
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   660
      Index           =   1
      Left            =   16140
      TabIndex        =   6
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
      Picture         =   "FrmFee.frx":F7F4
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   660
      Index           =   2
      Left            =   16920
      TabIndex        =   4
      Top             =   12210
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
      Picture         =   "FrmFee.frx":FB45
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   405
      Left            =   9450
      TabIndex        =   2
      Top             =   11790
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   714
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
      CalendarBackColor=   12648447
      CalendarForeColor=   12582912
      CalendarTitleBackColor=   8421504
      CalendarTitleForeColor=   12632256
      CalendarTrailingForeColor=   8421504
      Format          =   16646144
      CurrentDate     =   36927
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   405
      Left            =   12750
      TabIndex        =   3
      Top             =   11790
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   714
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
      CalendarBackColor=   12648447
      CalendarForeColor=   12582912
      CalendarTitleBackColor=   8421504
      CalendarTitleForeColor=   12632256
      CalendarTrailingForeColor=   8421504
      Format          =   16646144
      CurrentDate     =   36927
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
      Left            =   13140
      TabIndex        =   20
      Top             =   10770
      Width           =   900
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "����"
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
      Height          =   375
      Index           =   9
      Left            =   15270
      TabIndex        =   19
      Top             =   11835
      Width           =   720
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "����"
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
      Height          =   375
      Index           =   7
      Left            =   11970
      TabIndex        =   18
      Top             =   11835
      Width           =   720
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "��ȸ�Ⱓ :"
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
      Height          =   315
      Index           =   5
      Left            =   8040
      TabIndex        =   17
      Top             =   11835
      Width           =   1245
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '����
      Caption         =   "��       �� :"
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
      Left            =   3360
      TabIndex        =   13
      Top             =   12450
      Width           =   1305
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '����
      Caption         =   "������ȣ :"
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
      Left            =   3360
      TabIndex        =   12
      Top             =   11850
      Width           =   1305
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
      TabIndex        =   11
      Top             =   11670
      Width           =   2115
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H006F3C2F&
      BackStyle       =   1  '�������� ����
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00C0C0C0&
      Height          =   1755
      Left            =   240
      Top             =   11400
      Width           =   18735
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
      Left            =   8460
      TabIndex        =   16
      Top             =   10800
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
      Left            =   10530
      TabIndex        =   15
      Top             =   10800
      Width           =   1425
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
      TabIndex        =   14
      Top             =   1830
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '����
      Caption         =   "������� ���� ���� "
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
      TabIndex        =   10
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
      Left            =   14250
      TabIndex        =   9
      Top             =   10740
      Width           =   2025
   End
End
Attribute VB_Name = "FrmFee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim i As Integer

Left = 0
Top = 0
'Left = (Screen.Width - Width) / 2   ' ���� ���η� �߾ӿ� �����ϴ�.
'Top = (Screen.Height - Height) / 2   ' ���� ���η� �߾ӿ� �����ϴ�.
'Me.cmb_Gubun = Me.cmb_Gubun.List(0)

DTPicker1.value = Now
DTPicker2.value = Now

Glo_SQL_REG = "SELECT * FROM TB_FEE WHERE (REG_DATE >= '" & Format(DTPicker1, "yyyy-mm-dd") & " 00:00:00') AND (REG_DATE <= '" & Format(DTPicker2, "yyyy-mm-dd") & " 23:59:59') ORDER BY REG_DATE ASC"
'Glo_SQL_REG = "SELECT * From TB_FEE WHERE ORDER BY REG_DATE ASC"

Call ListView_REG_Draw
Call ListView_REG_SQL

List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    ������� �������� ����...!!", 0
Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    ������� �������� ����...!!")

End Sub

Public Sub ListView_REG_SQL()
Dim rs As Recordset
Dim Qry As String
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
    itmX.SubItems(1) = "" & rs!CAR_NO
    itmX.SubItems(2) = "" & rs!CAR_MODEL
    itmX.SubItems(3) = "" & rs!CAR_GUBUN
    itmX.SubItems(4) = "" & rs!CAR_FEE
    itmX.SubItems(5) = "" & rs!DRIVER_NAME
    itmX.SubItems(6) = "" & rs!DRIVER_PHONE
    itmX.SubItems(7) = "" & rs!DRIVER_DEPT
    itmX.SubItems(8) = "" & rs!DRIVER_CLASS
    itmX.SubItems(9) = "" & rs!Start_Date
    itmX.SubItems(10) = "" & rs!End_Date
    itmX.SubItems(11) = "" & rs!REG_DATE
    TOTAL_FEE = TOTAL_FEE + Val(rs!CAR_FEE)
    lbl_Search.Caption = TOTAL_FEE & "��"
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
    .ListView_REG.ColumnHeaders.Add , , " ������ȣ        "
    .ListView_REG.ColumnHeaders.Add , , " ������     "
    .ListView_REG.ColumnHeaders.Add , , " ��������   "
    .ListView_REG.ColumnHeaders.Add , , " �������   "
    .ListView_REG.ColumnHeaders.Add , , " ��    ��     "
    .ListView_REG.ColumnHeaders.Add , , " �� �� ó              "
    .ListView_REG.ColumnHeaders.Add , , " ��    ��    "
    .ListView_REG.ColumnHeaders.Add , , " ��    ��    "
    .ListView_REG.ColumnHeaders.Add , , " �� �� ��      "
    .ListView_REG.ColumnHeaders.Add , , " �� �� ��      "
    .ListView_REG.ColumnHeaders.Add , , " �� �� ��               "
    .ListView_REG.ColumnHeaders.Add , , " "
    
    For Column_to_size = 0 To .ListView_REG.ColumnHeaders.Count - 2
         SendMessage .ListView_REG.hwnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
    Next
End With
End Sub

Private Sub ListView_REG_ItemClick(ByVal Item As ComctlLib.ListItem)
ListView_REG.SetFocus
'txt_CarNo = ListView_REG.SelectedItem.SubItems(1)
End Sub

Private Sub cmd_Button_Click(Index As Integer)
Dim myExcelFile As New ExcelFile
Dim tmpFileName As String
Dim sql_str As String

Select Case Index
    Case 0  '����
        List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    ������� �������� ����", 0
        Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    ������� �������� ����")
        Unload Me
        Exit Sub

    Case 1
        tmpFileName = Format(Now, "YYYYMMDD_HHMMSS")
        tmpFileName = App.Path & "\Excel\" & tmpFileName & "_�������_��������" & ".xls"
        Call makeexcel(ListView_REG, tmpFileName, "�������_��������")
        Exit Sub
        
    Case 2
        '������������ �˻�

        
        Me.MousePointer = 11
        
        Glo_SQL_SEARCH = ""
        '���� ����
        '��ȸ�Ⱓ
        'Glo_SQL_REG = "SELECT * FROM TB_FEE WHERE (REG_DATE >= '" & Format(DTPicker1, "yyyy-mm-dd") & " 00:00:00') AND (REG_DATE <= '" & Format(DTPicker2, "yyyy-mm-dd") & " 23:59:59') ORDER BY REG_DATE ASC"
        sql_str = "SELECT * FROM TB_FEE WHERE (REG_DATE >= '" & Format(DTPicker1, "yyyy-mm-dd") & " 00:00:00') AND (REG_DATE <= '" & Format(DTPicker2, "yyyy-mm-dd") & " 23:59:59')"
        '������ȣ �˻�
        If (txt_CarNo.Text <> "") Then
            sql_str = sql_str & " AND (CAR_NO Like '%" & txt_CarNo.Text & "%')"
        End If
        '������ �̸� �˻�
        If (txt_Name.Text <> "") Then
            sql_str = sql_str & " AND (DRIVER_NAME Like '%" & txt_Name.Text & "%')"
        End If
        sql_str = sql_str & " ORDER BY REG_DATE ASC"
        Glo_SQL_REG = sql_str
        'Debug.Print Glo_SQL_REG
        Call ListView_REG_Draw
        Call ListView_REG_SQL
        Me.MousePointer = 0
        Exit Sub

End Select

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
