VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCardResult 
   Caption         =   " �ſ�ī�� ���系��"
   ClientHeight    =   11505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15345
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   11505
   ScaleWidth      =   15345
   StartUpPosition =   3  'Windows �⺻��
   Begin VB.CommandButton btn_Exit 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   14250
      TabIndex        =   16
      Top             =   90
      Width           =   1035
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
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   9600
      TabIndex        =   1
      Top             =   1560
      Width           =   1335
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
      Left            =   30
      TabIndex        =   0
      Top             =   10770
      Width           =   15315
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   7245
      Left            =   30
      TabIndex        =   2
      Top             =   2160
      Width           =   15270
      _ExtentX        =   26935
      _ExtentY        =   12779
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
      Height          =   510
      Index           =   1
      Left            =   11340
      TabIndex        =   3
      Top             =   1500
      Width           =   1230
      _Version        =   65536
      _ExtentX        =   2170
      _ExtentY        =   900
      _StockProps     =   78
      Caption         =   "Excel"
      ForeColor       =   14737632
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
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   510
      Index           =   2
      Left            =   12720
      TabIndex        =   4
      Top             =   1500
      Width           =   1230
      _Version        =   65536
      _ExtentX        =   2170
      _ExtentY        =   900
      _StockProps     =   78
      Caption         =   "�� ��"
      ForeColor       =   14737632
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
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   1560
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
      Format          =   74907648
      CurrentDate     =   36927
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   5490
      TabIndex        =   6
      Top             =   1560
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
      Format          =   74907648
      CurrentDate     =   36927
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "������ȣ :"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   11.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   8520
      TabIndex        =   15
      Top             =   1635
      Width           =   975
   End
   Begin VB.Label lbl_Search 
      Alignment       =   1  '������ ����
      BackStyle       =   0  '����
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   15.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1485
      TabIndex        =   14
      Top             =   10080
      Width           =   3165
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '����
      Caption         =   "�ſ�ī�� �������� "
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
      Left            =   390
      TabIndex        =   13
      Top             =   840
      Width           =   3525
   End
   Begin VB.Label lbl_COUNT 
      BackStyle       =   0  '����
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2445
      TabIndex        =   12
      Top             =   9690
      Width           =   1425
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '����
      Caption         =   "��ȸ�� �����Ǽ� :"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   0
      Left            =   375
      TabIndex        =   11
      Top             =   9690
      Width           =   1875
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "��ȸ�Ⱓ :"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   11.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   1530
      TabIndex        =   10
      Top             =   1635
      Width           =   975
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   11.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   7
      Left            =   4770
      TabIndex        =   9
      Top             =   1635
      Width           =   420
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   11.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   9
      Left            =   7650
      TabIndex        =   8
      Top             =   1635
      Width           =   420
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "Total : "
      BeginProperty Font 
         Name            =   "�������"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   375
      TabIndex        =   7
      Top             =   10110
      Width           =   900
   End
End
Attribute VB_Name = "frmCardResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public CardQry As String


Private Sub btn_Exit_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim i As Integer

    'cmb_DCGubun
'    Qry = "SELECT tb_calculate.DC_GUBUN From tb_calculate Group By tb_calculate.DC_GUBUN "
'    Set rs = New ADODB.Recordset
'    rs.Open Qry, adoConn
'    Do While Not (rs.EOF)
'        cmb_DCGubun.AddItem rs!DC_Gubun
'        rs.MoveNext
'    Loop
'    Set rs = Nothing
    
    
Left = (Screen.Width - Width) / 2   ' ���� ���η� �߾ӿ� �����ϴ�.
Top = (Screen.Height - Height) / 2   ' ���� ���η� �߾ӿ� �����ϴ�.
'Left = 0
'Top = 0
'Me.cmb_Gubun = Me.cmb_Gubun.List(0)
DTPicker1.value = Now
DTPicker2.value = Now
'DTPicker3.value = Format("00:00:00")
'DTPicker4.value = Format("23:59:59")

CardQry = "SELECT * FROM tb_kicc_log WHERE (Reg_Date >= '" & Format(DTPicker1, "yyyy-mm-dd") & " 00:00:00') AND (Reg_Date <= '" & Format(DTPicker2, "yyyy-mm-dd") & " 23:59:59') ORDER BY Reg_Date"

Call ListView1_Draw
Call ListView1_SQL
Call DataLogger("ī��������� ����...!!")

End Sub

Public Sub ListView1_SQL()
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
rs.Open CardQry, adoConn
lbl_COUNT = rs.RecordCount
Do While Not (rs.EOF)
    Set itmX = ListView1.ListItems.Add(, , "" & INDEX_NO)
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
    Call ListViewExtended(.ListView1)
    .ListView1.View = lvwReport
    .ListView1.ListItems.Clear
    .ListView1.ColumnHeaders.Clear
    .ListView1.ColumnHeaders.Add , , " No  "
    .ListView1.ColumnHeaders.Add , , " TicketCode                 "
    .ListView1.ColumnHeaders.Add , , " �����Ͻ�           "
    .ListView1.ColumnHeaders.Add , , " ī������                "
    .ListView1.ColumnHeaders.Add , , " ī���           "
    .ListView1.ColumnHeaders.Add , , " ����ݾ�          "
    .ListView1.ColumnHeaders.Add , , " ������ȣ          "
    .ListView1.ColumnHeaders.Add , , " RegDate                   "
    '.ListView1.ColumnHeaders.Add , , " "
    For Column_to_size = 0 To .ListView1.ColumnHeaders.Count - 2
         SendMessage .ListView1.hwnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
    Next
End With
End Sub

'Private Sub ListView_REG_ItemClick(ByVal Item As ComctlLib.ListItem)
'Dim Tmp_Path1, Tmp_Path2 As String
'
'ListView_REG.SetFocus
'
'lbl_Name(5) = ListView_REG.SelectedItem.SubItems(1) & " " & ListView_REG.SelectedItem.SubItems(2)
'lbl_Name(6) = ListView_REG.SelectedItem.SubItems(3) & " " & ListView_REG.SelectedItem.SubItems(4)
'lbl_Name(7) = ListView_REG.SelectedItem.SubItems(13) & " ��"
'lbl_Name(8) = ListView_REG.SelectedItem.SubItems(6) & " ��"
'lbl_Name(9) = ListView_REG.SelectedItem.SubItems(7) & " ��"
'lbl_Name(10) = ListView_REG.SelectedItem.SubItems(7) & " ��"
'
'End Sub

Private Sub cmd_Button_Click(Index As Integer)
Dim myExcelFile As New ExcelFile
Dim tmpFileName As String
Dim sql_str As String

Select Case Index
    Case 0  '����
        Call DataLogger("ī��������� ����")
        Unload Me
        Exit Sub

    Case 1
        'tmpFileName = Format(Now, "YYYYMMDD_HHMMSS")
        'tmpFileName = App.Path & "\Excel\" & tmpFileName & "_��������" & ".xls"
        'Call makeexcel(ListView_REG, tmpFileName, "_��������")
        tmpFileName = Format(Now, "YYYYMMDD_HHMMSS")
        tmpFileName = App.Path & "\Excel\" & tmpFileName & "_ī����系��.xls"
        'Call MakeCSV(ListView1, tmpFileName)
        Call makeexcel(ListView1, tmpFileName, "ī��������� ���� : " & lbl_Search.Caption)
        Exit Sub
        
    Case 2
        '������������ �˻�
        Me.MousePointer = 11
        CardQry = ""
        '���� ����
        '��ȸ�Ⱓ
        sql_str = "SELECT * FROM tb_kicc_log WHERE ( Reg_Date >= '" & Format(DTPicker1, "yyyy-mm-dd") & " 00:00:00') AND ( Reg_Date <= '" & Format(DTPicker2, "yyyy-mm-dd") & " 23:59:59')"
        '������ȣ �˻�
        If (txt_CarNo.Text <> "") Then
            If IsNumeric(txt_CarNo) And Len(txt_CarNo) = 4 Then
            Else
                MsgBox "������ȣ ��4�ڸ��� Ȯ�����ּ���."
                Me.MousePointer = 0
                Exit Sub
            End If
            sql_str = sql_str & " AND (CarNum Like '%" & txt_CarNo.Text & "')"
        End If
        
'        If (Len(cmb_DCGubun.Text) <> 0) Then
'            sql_str = sql_str & " AND (DC_GUBUN Like '" & cmb_DCGubun & "')"
'        End If
        
        sql_str = sql_str & " ORDER BY Reg_Date"
        CardQry = sql_str
        'List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & sql_str, 0
        Call DataLogger(sql_str)
        
        'Debug.Print CardQry
        Call ListView1_Draw
        Call ListView1_SQL
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



