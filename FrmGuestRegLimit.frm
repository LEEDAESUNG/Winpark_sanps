VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMctl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmGuestRegLimit 
   BackColor       =   &H00404040&
   BorderStyle     =   1  '���� ����
   Caption         =   "ParkingManager��"
   ClientHeight    =   9435
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13245
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9435
   ScaleWidth      =   13245
   Begin VB.Frame Frame2 
      BackColor       =   &H00404040&
      Caption         =   " �������� ����"
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
      Height          =   2625
      Left            =   135
      TabIndex        =   12
      Top             =   6720
      Width           =   12975
      Begin VB.TextBox txt_etc 
         BeginProperty Font 
            Name            =   "�������"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         IMEMode         =   10  '�ѱ� 
         Left            =   5730
         TabIndex        =   3
         Text            =   "��Ÿ�����Է�"
         Top             =   690
         Width           =   6990
      End
      Begin VB.TextBox txt_Carno 
         BeginProperty Font 
            Name            =   "�������"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         IMEMode         =   10  '�ѱ� 
         Left            =   1995
         TabIndex        =   1
         Text            =   "����12��3456"
         Top             =   690
         Width           =   2505
      End
      Begin VB.TextBox txt_MaxInParkCount 
         BeginProperty Font 
            Name            =   "�������"
            Size            =   14.25
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1995
         TabIndex        =   2
         Text            =   "10"
         Top             =   1260
         Visible         =   0   'False
         Width           =   2505
      End
      Begin Threed.SSCommand SSCommand3 
         Height          =   690
         Left            =   7140
         TabIndex        =   6
         Top             =   1755
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   1217
         _StockProps     =   78
         Caption         =   "�ʱ�ȭ"
         ForeColor       =   16777215
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
         Picture         =   "FrmGuestRegLimit.frx":0000
      End
      Begin Threed.SSCommand SSCommand4 
         Height          =   690
         Left            =   8550
         TabIndex        =   4
         Top             =   1755
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   1217
         _StockProps     =   78
         Caption         =   "���"
         ForeColor       =   16777215
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
         Picture         =   "FrmGuestRegLimit.frx":0351
      End
      Begin Threed.SSCommand SSCommand5 
         Height          =   690
         Left            =   9960
         TabIndex        =   8
         Top             =   1755
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   1217
         _StockProps     =   78
         Caption         =   "����"
         ForeColor       =   16777215
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
         Picture         =   "FrmGuestRegLimit.frx":06A2
      End
      Begin Threed.SSCommand SSCommand6 
         Height          =   690
         Left            =   11370
         TabIndex        =   9
         Top             =   1755
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   1217
         _StockProps     =   78
         Caption         =   "����"
         ForeColor       =   16777215
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
         Picture         =   "FrmGuestRegLimit.frx":09F3
      End
      Begin Threed.SSCommand cmd_Search 
         Height          =   690
         Left            =   5730
         TabIndex        =   7
         Top             =   1755
         Width           =   1425
         _Version        =   65536
         _ExtentX        =   2514
         _ExtentY        =   1217
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
         Picture         =   "FrmGuestRegLimit.frx":0D44
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "�޸�"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   5115
         TabIndex        =   15
         Top             =   750
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "������ȣ"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   855
         TabIndex        =   14
         Top             =   750
         Width           =   1080
      End
      Begin VB.Label lbl_dong 
         AutoSize        =   -1  'True
         BackStyle       =   0  '����
         Caption         =   "�ִ�湮Ƚ��"
         BeginProperty Font 
            Name            =   "�������"
            Size            =   14.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   315
         TabIndex        =   13
         Top             =   1335
         Visible         =   0   'False
         Width           =   1620
      End
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   5100
      Left            =   135
      TabIndex        =   0
      Top             =   1185
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   8996
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9150
      Top             =   150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Threed.SSCommand SSCommand2 
      Height          =   570
      Left            =   11835
      TabIndex        =   5
      Top             =   120
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   1005
      _StockProps     =   78
      Caption         =   "�ݱ�"
      ForeColor       =   16777215
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
      Picture         =   "FrmGuestRegLimit.frx":1095
   End
   Begin Threed.SSCommand SSCommand7 
      Height          =   570
      Left            =   10470
      TabIndex        =   10
      Top             =   120
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   1005
      _StockProps     =   78
      Caption         =   "����"
      ForeColor       =   16777215
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
      Picture         =   "FrmGuestRegLimit.frx":13E6
   End
   Begin Threed.SSPanel PnlOut 
      Height          =   390
      Index           =   7
      Left            =   10605
      TabIndex        =   16
      Top             =   6315
      Width           =   2520
      _Version        =   65536
      _ExtentX        =   4445
      _ExtentY        =   688
      _StockProps     =   15
      Caption         =   "  �˻� �Ǽ�"
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
         TabIndex        =   17
         Top             =   60
         Width           =   1275
      End
   End
   Begin Threed.SSPanel PnlOut 
      Height          =   390
      Index           =   0
      Left            =   8415
      TabIndex        =   18
      Top             =   795
      Width           =   4755
      _Version        =   65536
      _ExtentX        =   8387
      _ExtentY        =   688
      _StockProps     =   15
      Caption         =   " �湮Ƚ�� : 1��~���ϱ���, �ſ�1�� 0ȸ�� �ʱ�ȭ �˴ϴ�."
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
      Alignment       =   1
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   0
      X1              =   135
      X2              =   13110
      Y1              =   780
      Y2              =   765
   End
   Begin VB.Label lbl_APS 
      BackStyle       =   0  '����
      Caption         =   "�湮�������� ��ȸ/���"
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
      Left            =   165
      TabIndex        =   11
      Top             =   300
      Width           =   4470
   End
End
Attribute VB_Name = "FrmGuestRegLimit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nSeq As Long


Private Sub Form_Load()

    Dim sQry As String
    
    Left = (Screen.width - width) / 2   ' ���� ���η� �߾ӿ� �����ϴ�.
    Top = (Screen.height - height) / 2   ' ���� ���η� �߾ӿ� �����ϴ�.
    
    sQry = "SELECT * From tb_guest_limit ORDER BY CAR_NO "

    Call Clear_Field
    Call ListView1_Draw
    Call ListView1_SQL(sQry)
    
End Sub


Private Sub SSCommand2_Click()
    Unload Me
    'Me.Hide
End Sub

Private Function Get_MaxParkDay()
    
    Dim rs As Recordset
    Dim bQryResult As Boolean
    Dim nMaxParkDay As Integer
    
    nMaxParkDay = 0
    
    If (Glo_GuestReg_YN = "Y") Then
        Set rs = New ADODB.Recordset: bQryResult = DataBaseQuery(rs, adoConn, "SELECT * FROM TB_CONFIG WHERE NAME = 'GuestCarReg_MaxParkDay' ", False): nMaxParkDay = rs!Content: Set rs = Nothing
    End If
    
    Get_MaxParkDay = nMaxParkDay
    
End Function
Private Sub Clear_Field()
    
    SSCommand4.Enabled = True '��Ϲ�ư
    SSCommand5.Enabled = False '������ư
    SSCommand6.Enabled = False '������ư
    
    txt_Carno = ""
    txt_MaxInParkCount = "10"
    txt_etc = ""

End Sub
Private Sub ListView1_Draw()
    Dim Column_to_size As Integer

'On Error GoTo Err_p
On Error Resume Next

    Call ListViewExtended(ListView1)
    ListView1.View = lvwReport
    ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , " ��Ϲ�ȣ "
    ListView1.ColumnHeaders.Add , , " ������ȣ          "
'    ListView1.ColumnHeaders.Add , , " �ִ�湮Ƚ��(��)  "
'    ListView1.ColumnHeaders.Add , , " �湮Ƚ��  "
    ListView1.ColumnHeaders.Add , , " �޸�                                                                        "
    ListView1.ColumnHeaders.Add , , " ��ϳ�¥      "
    
    
    For Column_to_size = 0 To ListView1.ColumnHeaders.Count - 1
         SendMessage ListView1.hwnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
    Next
End Sub

Private Sub ListView1_SQL(qry As String)
    Dim bQryResult As Boolean
    Dim rs As Recordset
    Dim itmX As ListItem
    Dim INDEX_NO As Long
    Dim i As Integer
    
    Set rs = New ADODB.Recordset
    bQryResult = DataBaseQuery(rs, adoConn, qry, False)
    If (bQryResult = False) Then
        Exit Sub
    End If

    LblRecordCount = "0"
    INDEX_NO = 0
    Do While Not (rs.EOF)
        Set itmX = ListView1.ListItems.Add(, , "" & rs!SEQ)
        i = 1
        itmX.SubItems(i) = "" & rs!CAR_NO: i = i + 1
'        itmX.SubItems(i) = "" & rs!MAXINPARK: i = i + 1
'        itmX.SubItems(i) = "" & rs!NOWINPARK: i = i + 1
        itmX.SubItems(i) = "" & rs!TEMP1: i = i + 1
        itmX.SubItems(i) = "" & Format(rs!REG_DATE, "yyyy-mm-dd hh:nn:ss"): i = i + 1
        INDEX_NO = INDEX_NO + 1
        
        rs.MoveNext
    Loop
    Set rs = Nothing
    
    LblRecordCount.Caption = INDEX_NO
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
    On Error Resume Next
    
    ListView1.SetFocus

    nSeq = ListView1.SelectedItem
    
    txt_Carno = ListView1.SelectedItem.SubItems(1)
    'txt_MaxInParkCount = ListView1.SelectedItem.SubItems(2)
    'txt_etc = ListView1.SelectedItem.SubItems(4)
    txt_etc = ListView1.SelectedItem.SubItems(2)

    SSCommand4.Enabled = False '��Ϲ�ư
    SSCommand5.Enabled = True '������ư
    SSCommand6.Enabled = True '������ư
    
End Sub

'�ʱ�ȭ
Private Sub SSCommand3_Click()
    Call Clear_Field
End Sub

'���
Private Sub SSCommand4_Click()

    If (Check_Field = False) Then
        Msg_Box.Label1 = "�湮�������� �Է� �����Դϴ�" & vbCrLf & vbCrLf & "���Է� �ٶ��ϴ�"
        Msg_Box.Show 1
        Exit Sub
    End If
    
    If (isNewRecord = False) Then
        Msg_Box.Label1 = "�̹� ��ϵ� �湮�������� �Դϴ�" & vbCrLf & vbCrLf & "���Է� �ٶ��ϴ�"
        Msg_Box.Show 1
        Exit Sub
    End If
    
    
    Dim sLog As String
    Dim sNowDT As String
    
    sNowDT = Format(Now, "yyyy-mm-dd hh:nn:ss")
    
    sLog = "�湮�������� ���:" & Glo_Login_ID & ":" & txt_Carno & ""
    adoConn.Execute "insert into tb_guest_limit (CAR_NO,MAXINPARK,NOWINPARK,UPDATE_DATE,REG_DATE,TEMP1) VALUES ( '" & txt_Carno & "','" & txt_MaxInParkCount & "','0','" & sNowDT & "','" & sNowDT & "', '" & txt_etc & "') "
    adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('�湮����', '" & txt_Carno & "', '" & sLog & "', '" & txt_MaxInParkCount & "', " & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
    Call DataLogger(sLog)

    Call Clear_Field
    Call ListView1_Draw
    Call ListView1_SQL("SELECT * From tb_guest_limit order by car_no")
End Sub

Private Function Check_Field() As Boolean
    Dim bCheck As Boolean
    bCheck = True
    
    If Not ((LenH(txt_Carno.text) = 11) Or (LenH(txt_Carno.text) = 12) Or (LenH(txt_Carno.text) = 8) Or (LenH(txt_Carno.text) = 9)) Then
        txt_Carno = "":             txt_Carno.SetFocus
        bCheck = False
    End If
    
    If (IsNumeric(txt_MaxInParkCount) = False) Then
        txt_MaxInParkCount = "":    txt_MaxInParkCount.SetFocus
        bCheck = False
    End If
    
    Check_Field = bCheck
End Function

Private Function isNewRecord() As Boolean
    Dim rs As Recordset
    Dim bisNew As Boolean
    
    bisNew = False

    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM tb_guest_limit WHERE CAR_NO = '" & txt_Carno & "' ", adoConn
    If rs.EOF Then
        bisNew = True
    End If
    Set rs = Nothing
    
    isNewRecord = bisNew
End Function




'����
Private Sub SSCommand5_Click()
    If (nSeq < 0) Then
        Msg_Box.Label1 = "������ ������ �����ϼ���"
        Msg_Box.Show 1
        Exit Sub
    End If
    
    If (Check_Field = False) Then
        Msg_Box.Label1 = "�湮�������� ���� �����Դϴ�" & vbCrLf & vbCrLf & "���Է� �ٶ��ϴ�"
        Msg_Box.Show 1
        Exit Sub
    End If
    
'    If (isNewRecord = False) Then
'        Msg_Box.Label1 = "�湮���� ��ϵ� ������ȣ�Դϴ�" & vbCrLf & vbCrLf & "���Է� �ٶ��ϴ�"
'        Msg_Box.Show 1
'        Exit Sub
'    End If
    
    MBox.Label3.Caption = txt_Carno.text
    MBox.Label1.Caption = "�湮�������� ������ �����մϴ�." & vbCrLf & " �����Ͻðڽ��ϱ�?"
    MBox.Label2.Caption = "�湮�������� ����"
    MBox.Show 1
    If (Glo_MsgRet = True) Then
        
        Dim sLog As String
        
        sLog = "�湮�������� ����:" & Glo_Login_ID & ":" & txt_Carno
        adoConn.Execute "UPDATE tb_guest_limit SET CAR_NO = '" & txt_Carno & "', MAXINPARK= '" & txt_MaxInParkCount & "', TEMP1 = '" & txt_etc & "' WHERE SEQ = '" & nSeq & "' "
        adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('�湮����', '" & txt_Carno & "', '" & sLog & "', '" & txt_MaxInParkCount & "', " & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
        Call DataLogger(sLog)
        
        Call Clear_Field
        Call ListView1_Draw
        Call ListView1_SQL("SELECT * From tb_guest_limit order by car_no")
        
    End If
End Sub

'����
Private Sub SSCommand6_Click()
    If (nSeq < 0) Then
        Msg_Box.Label1 = "������ ������ �����ϼ���"
        Msg_Box.Show 1
        Exit Sub
    End If
    
    
    MBox.Label3.Caption = txt_Carno.text
    MBox.Label1.Caption = "�湮�������� ������ �����մϴ�." & vbCrLf & " �����Ͻðڽ��ϱ�?"
    MBox.Label2.Caption = "�湮�������� ����"
    MBox.Show 1
    If (Glo_MsgRet = True) Then
        
        Dim sLog As String
        sLog = "�湮�������� ����:" & Glo_Login_ID & ":" & txt_Carno
        adoConn.Execute "DELETE FROM tb_guest_limit WHERE SEQ = '" & nSeq & "' "
        adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('�湮����', '" & txt_Carno & "', '" & sLog & "', '" & txt_MaxInParkCount & "', " & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
        Call DataLogger(sLog)
        
        Call Clear_Field
        Call ListView1_Draw
        Call ListView1_SQL("SELECT * From tb_guest_limit order by car_no")
    End If
End Sub

'�˻�
Private Sub cmd_Search_Click()
    Dim sQry As String

    sQry = "SELECT * From tb_guest_limit "

    If (Len(txt_Carno) > 0) Then
        sQry = sQry & " WHERE CAR_NO LIKE '%" & txt_Carno & "%' "
    End If
    sQry = sQry & " ORDER BY CAR_NO "

    Call Clear_Field
    Call ListView1_Draw
    Call ListView1_SQL(sQry)
End Sub



Private Sub SSCommand7_Click()

    Dim tmpFileName As String
On Error GoTo Err_P
    tmpFileName = Format(Now, "YYYYMMDD_HHMMSS")
    tmpFileName = App.Path & "\Excel\" & tmpFileName & "_�湮�������� ��ϳ���"
    
    CommonDialog1.CancelError = True
    CommonDialog1.InitDir = App.Path
    CommonDialog1.Filter = "��������(*.csv)|*.csv"
    CommonDialog1.fileName = tmpFileName
    CommonDialog1.ShowSave
    tmpFileName = CommonDialog1.fileName
    tmpFileName = Mid(tmpFileName, 1, Len(tmpFileName) - 4)

    Call MakeCSV(ListView1, tmpFileName)
    Exit Sub
Err_P:
     Select Case Err
    Case 32755 '  Dialog Cancelled
        'MsgBox "you cancelled the dialog box"
    Case Else
        'MsgBox "Unexpected error. Err " & Err & " : " & Error
    End Select
End Sub



Private Sub txt_Carno_LostFocus()
    If Not ((LenH(txt_Carno.text) = 11) Or (LenH(txt_Carno.text) = 12) Or (LenH(txt_Carno.text) = 8) Or (LenH(txt_Carno.text) = 9)) Then
    End If
End Sub

Private Sub txt_MaxInParkCount_KeyPress(KeyAscii As Integer)
    
    '�������Է�
    If (txt_MaxInParkCount = "0") Then
        txt_MaxInParkCount = ""
    End If

    If (KeyAscii = 45) Then ' -
        txt_MaxInParkCount = ""
    ElseIf (KeyAscii = vbKeyBack Or (KeyAscii >= vbKey0 And KeyAscii <= vbKey9)) Then '�齺���̽�, ����
    Else
        KeyAscii = 0
    End If
    
    txt_MaxInParkCount = Format(txt_MaxInParkCount, "#####0")
End Sub
