VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMctl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmGuestRegLog 
   BackColor       =   &H00404040&
   BorderStyle     =   1  '���� ����
   Caption         =   "ParkingManager��"
   ClientHeight    =   11715
   ClientLeft      =   5160
   ClientTop       =   1725
   ClientWidth     =   17670
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   11715
   ScaleWidth      =   17670
   Begin VB.Timer Timer_CheckSignup 
      Interval        =   10000
      Left            =   6240
      Top             =   345
   End
   Begin VB.TextBox txt_CarNo 
      BeginProperty Font 
         Name            =   "�������"
         Size            =   20.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      IMEMode         =   10  '�ѱ� 
      Left            =   12750
      MaxLength       =   10
      TabIndex        =   2
      Top             =   2280
      Width           =   4260
   End
   Begin VB.ComboBox cmb_GuestDong 
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   13560
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   1
      Top             =   1695
      Width           =   1320
   End
   Begin VB.ComboBox cmb_GuestHo 
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   15660
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   0
      Top             =   1695
      Width           =   1320
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   7290
      Left            =   30
      TabIndex        =   3
      Top             =   4395
      Width           =   17610
      _ExtentX        =   31062
      _ExtentY        =   12859
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
   Begin Threed.SSPanel PnlOut 
      Height          =   390
      Index           =   7
      Left            =   14505
      TabIndex        =   4
      Top             =   3960
      Width           =   2520
      _Version        =   65536
      _ExtentX        =   4445
      _ExtentY        =   688
      _StockProps     =   15
      Caption         =   " �˻� �Ǽ�"
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
         TabIndex        =   5
         Top             =   75
         Width           =   1275
      End
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   345
      Left            =   13050
      TabIndex        =   13
      Top             =   1155
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   609
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
      Format          =   76414976
      CurrentDate     =   36927
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   345
      Left            =   15150
      TabIndex        =   14
      Top             =   1155
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   609
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
      Format          =   76414976
      CurrentDate     =   36927
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   570
      Left            =   14370
      TabIndex        =   15
      Top             =   225
      Width           =   1260
      _Version        =   65536
      _ExtentX        =   2222
      _ExtentY        =   1005
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
      Picture         =   "FrmGuestRegLog.frx":0000
   End
   Begin Threed.SSCommand SSCommand2 
      Cancel          =   -1  'True
      Height          =   570
      Left            =   15720
      TabIndex        =   16
      Top             =   225
      Width           =   1260
      _Version        =   65536
      _ExtentX        =   2222
      _ExtentY        =   1005
      _StockProps     =   78
      Caption         =   "�ݱ�"
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
      Picture         =   "FrmGuestRegLog.frx":0351
   End
   Begin Threed.SSCommand cmd_GuestRegSetup 
      Height          =   585
      Left            =   18180
      TabIndex        =   17
      Top             =   270
      Visible         =   0   'False
      Width           =   1695
      _Version        =   65536
      _ExtentX        =   2990
      _ExtentY        =   1032
      _StockProps     =   78
      Caption         =   "�ð�����"
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
      Enabled         =   0   'False
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "FrmGuestRegLog.frx":06A2
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   17040
      Top             =   315
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Threed.SSCommand SSCommand3 
      Height          =   570
      Left            =   9285
      TabIndex        =   18
      ToolTipText     =   "�湮���� ���� ���Խ�û�ڿ� ���Ͽ� ���Խ��� ó��. ���Խ��� ����� ������� ����� �۾��� ����˴ϴ�."
      Top             =   3135
      Width           =   1665
      _Version        =   65536
      _ExtentX        =   2937
      _ExtentY        =   1005
      _StockProps     =   78
      Caption         =   "���Խ���/����"
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
      Picture         =   "FrmGuestRegLog.frx":09F3
   End
   Begin Threed.SSCommand SSCommand4 
      Height          =   570
      Left            =   11025
      TabIndex        =   19
      ToolTipText     =   "�湮���� ���� ��� �� ��ȸ�մϴ�"
      Top             =   3135
      Width           =   1665
      _Version        =   65536
      _ExtentX        =   2937
      _ExtentY        =   1005
      _StockProps     =   78
      Caption         =   "�湮������"
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
      Picture         =   "FrmGuestRegLog.frx":11B8
   End
   Begin Threed.SSCommand SSCommand5 
      Height          =   570
      Left            =   14910
      TabIndex        =   20
      ToolTipText     =   "�����湮���� �������� ��ȣ���� �����ð��� ��ȸ�մϴ�."
      Top             =   3135
      Width           =   2100
      _Version        =   65536
      _ExtentX        =   3704
      _ExtentY        =   1005
      _StockProps     =   78
      Caption         =   "�����ð�(��ȣ����)"
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
      Picture         =   "FrmGuestRegLog.frx":197D
   End
   Begin Threed.SSCommand SSCommand6 
      Height          =   570
      Left            =   12750
      TabIndex        =   21
      ToolTipText     =   "�����湮���� �������� ������ �����ð��� ��ȸ�մϴ�."
      Top             =   3135
      Width           =   2100
      _Version        =   65536
      _ExtentX        =   3704
      _ExtentY        =   1005
      _StockProps     =   78
      Caption         =   "�����ð�(������)"
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
      Picture         =   "FrmGuestRegLog.frx":2142
   End
   Begin Threed.SSCommand SSCommand7 
      Height          =   570
      Left            =   7560
      TabIndex        =   22
      ToolTipText     =   "�湮���� ���� ���Խ�û�ڿ� ���Ͽ� ���Խ��� ó��. ���Խ��� ����� ������� ����� �۾��� ����˴ϴ�."
      Top             =   3135
      Width           =   1665
      _Version        =   65536
      _ExtentX        =   2937
      _ExtentY        =   1005
      _StockProps     =   78
      Caption         =   "�湮����"
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
      Picture         =   "FrmGuestRegLog.frx":2907
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '����
      DataField       =   "imgpath1"
      DataSource      =   "Adodc1"
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   255
      TabIndex        =   12
      Top             =   13410
      Width           =   14715
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "��ȸ�Ⱓ"
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
      Height          =   285
      Index           =   5
      Left            =   12015
      TabIndex        =   11
      Top             =   1185
      Width           =   900
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "~"
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
      Height          =   285
      Index           =   7
      Left            =   14940
      TabIndex        =   10
      Top             =   1185
      Width           =   150
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "������ȣ"
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
      Height          =   315
      Index           =   10
      Left            =   11580
      TabIndex        =   9
      Top             =   2400
      Width           =   1080
   End
   Begin VB.Label lbl_GuestReservation 
      BackStyle       =   0  '����
      Caption         =   " �����湮����"
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
      Height          =   300
      Left            =   180
      TabIndex        =   8
      Top             =   480
      Width           =   5385
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   0
      X1              =   150
      X2              =   17010
      Y1              =   930
      Y2              =   930
   End
   Begin VB.Label lbl_GuestDong 
      Alignment       =   1  '������ ����
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "ȸ���"
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
      Height          =   285
      Left            =   12720
      TabIndex        =   7
      Top             =   1725
      Width           =   675
   End
   Begin VB.Label lbl_GuestHo 
      Alignment       =   1  '������ ����
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�μ�"
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
      Height          =   285
      Left            =   15015
      TabIndex        =   6
      Top             =   1725
      Width           =   570
   End
End
Attribute VB_Name = "FrmGuestRegLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'�湮���� �����湮�ð�����
Private Sub cmd_GuestRegSetup_Click()
    Me.MousePointer = 0
    FrmGuestRegCar.Show 1
    Call DataLogger("[HOST Button]    " & "�湮���� �����湮 ���� ����")
End Sub

Private Sub Form_Load()
    Dim Record_Source As String
    Dim i As Integer
    
'On Error GoTo err_P
    
    Left = (Screen.width - width) / 2   ' ���� ���η� �߾ӿ� �����ϴ�.
    Top = (Screen.height - height) / 2   ' ���� ���η� �߾ӿ� �����ϴ�.
    
    
    DTPicker1.value = Now
    DTPicker2.value = Now
    
    If (Glo_User_Type = "����1/����2") Then
        lbl_GuestDong = "�Ҽ�"
        lbl_GuestHo = "����"
        SSCommand5.Caption = "�����ð�(�μ���)"
    Else
        lbl_GuestDong = "��"
        lbl_GuestHo = "ȣ"
        SSCommand5.Caption = "�����ð�(��ȣ����)"
    End If
    
    
    Call SetDong
    Call SetHo
    
Exit Sub
    
Err_p:
    MsgBox "������ ���̽� �������" & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & "��Ʈ�� �����ڿ��� ���� �ٶ��ϴ�." & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & "������ ���̽� ���������� �ڷ�˻� ����� �����Ҽ� �����ϴ�.", vbCritical
End Sub



Private Sub SetDong()
    Dim rs As Recordset
    Dim sQry As String
    
On Error GoTo Err_p

    cmb_GuestDong.Clear
    cmb_GuestDong.AddItem "��ü"
    
    sQry = "SELECT DRIVER_DEPT From tb_guestReg_admin Group By DRIVER_DEPT"
    Set rs = New ADODB.Recordset
    rs.Open sQry, adoConn
    
    If Not rs.EOF Then
        Do While Not (rs.EOF)
            cmb_GuestDong.AddItem "" & rs!DRIVER_DEPT
            rs.MoveNext
        Loop
    End If
    Set rs = Nothing

    cmb_GuestDong.ListIndex = 0
    
Exit Sub
Err_p:
    Call DataLogger("[FrmGuestRegLog SetDong]    " & Err.Description & " " & sQry)
End Sub

Private Sub SetHo()
    Dim rs As Recordset
    Dim sQry As String

On Error GoTo Err_p

    cmb_GuestHo.Clear
    cmb_GuestHo.AddItem "��ü"

    sQry = "SELECT DRIVER_CLASS From tb_guestReg_admin Group By DRIVER_CLASS"
    Set rs = New ADODB.Recordset
    rs.Open sQry, adoConn
    
    If Not rs.EOF Then
        Do While Not (rs.EOF)
            cmb_GuestHo.AddItem "" & rs!DRIVER_CLASS
            rs.MoveNext
        Loop
    End If
    Set rs = Nothing

    cmb_GuestHo.ListIndex = 0
    
Exit Sub
Err_p:
    Call DataLogger("[FrmGuestRegLog SetHo]    " & Err.Description & " " & sQry)
End Sub



Private Sub SetGuestDong()
    Dim bQryResult As Boolean
    Dim rs As Recordset
    Dim qry As String
    
    cmb_GuestDong.Clear
    cmb_GuestDong.AddItem "��ü"
    
    qry = "SELECT DONG From tb_guest_log Group By DONG"
    'QRY = "SELECT DONG From tb_guest_log Where tb_guest_log.DT_IN >= '" & Format(DTPicker5, "yyyy-mm-dd") & " 00:00:00' Group By DONG"
    Set rs = New ADODB.Recordset
    bQryResult = DataBaseQuery(rs, adoConn, qry, False)

    If Not rs.EOF Then
        Do While Not (rs.EOF)
            cmb_GuestDong.AddItem "" & rs!Dong
            rs.MoveNext
        Loop
    End If
    Set rs = Nothing
    
    cmb_GuestDong.ListIndex = 0
End Sub

Private Sub SetGuestHo()
    Dim bQryResult As Boolean
    Dim rs As Recordset
    Dim qry As String
    
    
    cmb_GuestHo.Clear
    cmb_GuestHo.AddItem "��ü"
    
    qry = "SELECT Ho From tb_guest_log Group By Ho"
    Set rs = New ADODB.Recordset
    bQryResult = DataBaseQuery(rs, adoConn, qry, False)
    
    If Not rs.EOF Then
        Do While Not (rs.EOF)
            cmb_GuestHo.AddItem "" & rs!Ho
            rs.MoveNext
        Loop
    End If
    Set rs = Nothing
    
    cmb_GuestHo.ListIndex = 0
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

Private Sub SSCommand1_Click()

    Dim tmpFileName As String
On Error GoTo Err_p
    tmpFileName = Format(Now, "YYYYMMDD_HHMMSS")
    tmpFileName = App.Path & "\Excel\" & tmpFileName & "_�����湮���������ð�"
        
        
    CommonDialog1.CancelError = True
    CommonDialog1.InitDir = App.Path
    CommonDialog1.Filter = "��������(*.csv)|*.csv"
    CommonDialog1.fileName = tmpFileName
    CommonDialog1.ShowSave
    tmpFileName = CommonDialog1.fileName
    tmpFileName = Mid(tmpFileName, 1, Len(tmpFileName) - 4)

    Call MakeCSV(ListView1, tmpFileName)
    Exit Sub
Err_p:
     Select Case Err
    Case 32755 '  Dialog Cancelled
        'MsgBox "you cancelled the dialog box"
    Case Else
        'MsgBox "Unexpected error. Err " & Err & " : " & Error
    End Select

End Sub

'����
Private Sub SSCommand2_Click()
    Unload Me
    'Me.Hide
End Sub

'�����Ҵ�ð� ���ϱ�
Private Function GetMaxParkTime(argDong As String, argHo As String)
    Dim rs As Recordset
    Dim sQry As String
    
    On Error GoTo Err_p
    
    sQry = "Select MAXPARKTIME AS MaxTime from tb_guestReg_admin "

    If (argDong <> "��ü") Then
        If (argHo <> "��ü") Then
            sQry = sQry & " WHERE DRIVER_DEPT = '" & argDong & "' AND DRIVER_CLASS = '" & argHo & "' "
        Else
            sQry = sQry & " WHERE DRIVER_DEPT = '" & argDong & "' "
        End If
    Else
        If (argHo <> "��ü") Then
            sQry = sQry & " WHERE DRIVER_CLASS = '" & argHo & "' "
        Else
        End If
    End If
    
    sQry = sQry & " ORDER BY MAXPARKTIME DESC"
    
    Set rs = New ADODB.Recordset
    rs.Open sQry, adoConn
    If Not (rs.EOF) Then
        If (IsNull(rs!MaxTime)) Then
            GetMaxParkTime = 0
        Else
            GetMaxParkTime = rs!MaxTime
        End If
    Else
        GetMaxParkTime = 0
    End If
    Set rs = Nothing
    
    Exit Function
    
Err_p:

    Set rs = Nothing
    GetMaxParkTime = 0
End Function

'����Ʈ�� ���
'��Ȯ�� ������ �ִ� ��츸 ó��
Public Sub ListView_Draw(sQry As String)
    
    Dim Column_to_size As Integer
    Dim rs As Recordset
    Dim INDEX_NO As Long
    Dim i As Integer
    
    
    Dim itmX As ListItem
    
    
    On Error GoTo Err_p
    
    Call ListViewExtended(ListView1)
    ListView1.View = lvwReport
    ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , " No "
    
    ListView1.ColumnHeaders.Add , , " ������ȣ         "
    If (Glo_User_Type = "����1/����2") Then
        ListView1.ColumnHeaders.Add , , " ��   ��1 "
        ListView1.ColumnHeaders.Add , , " ��   ��2 "
    Else
        ListView1.ColumnHeaders.Add , , " ��       "
        ListView1.ColumnHeaders.Add , , " ȣ       "
    End If
    ListView1.ColumnHeaders.Add , , " �湮��       "
    ListView1.ColumnHeaders.Add , , " ����ó                 "
    
    ListView1.ColumnHeaders.Add , , " ����Ͻ�                     "
    'ListView1.ColumnHeaders.Add , , " ��������                     "
    
    For Column_to_size = 0 To ListView1.ColumnHeaders.Count - 1
         SendMessage ListView1.hwnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
    Next
    
    

    INDEX_NO = 1
    
    Set rs = New ADODB.Recordset
    rs.Open sQry, adoConn
    Do While Not (rs.EOF)

        Set itmX = ListView1.ListItems.Add(, , "" & INDEX_NO)
        
        i = 1
        itmX.SubItems(i) = "" & rs!CAR_NO: i = i + 1
        itmX.SubItems(i) = "" & rs!DRIVER_DEPT: i = i + 1
        itmX.SubItems(i) = "" & rs!DRIVER_CLASS: i = i + 1
        itmX.SubItems(i) = "" & rs!DRIVER_NAME: i = i + 1
        itmX.SubItems(i) = "" & rs!DRIVER_PHONE: i = i + 1
        itmX.SubItems(i) = "" & rs!REG_DATE: i = i + 1
        'itmX.SubItems(i) = "" & rs!Pass_YN: i = i + 1
                
        INDEX_NO = INDEX_NO + 1

        rs.MoveNext
    Loop

    Set rs = Nothing

    LblRecordCount.Caption = INDEX_NO - 1
'    LblTotalParkTime = totalParkTime
    
    Exit Sub
Err_p:
    Set rs = Nothing
    Call DataLogger("FrmGuestRegLog ListView_Draw Err " & Err.Description & " " & sQry)
End Sub


Private Sub Draw_Listview_Guest(guest As stGuest, ByVal IndexNo As Integer, nMaxTime As Long)

    Dim itmX As ListItem
    Dim col As Integer
    Dim nOverTime As Long
    Dim nRemainTime As Long
    
    Set itmX = ListView1.ListItems.Add(, , "" & IndexNo)

    col = 1
    itmX.SubItems(col) = "" & guest.CarGubun: col = col + 1
    itmX.SubItems(col) = "" & guest.InCarNo: col = col + 1
    itmX.SubItems(col) = "" & guest.Dong: col = col + 1
    itmX.SubItems(col) = "" & guest.Ho: col = col + 1
    itmX.SubItems(col) = "" & guest.GuestName: col = col + 1
    itmX.SubItems(col) = "" & guest.Tel: col = col + 1
    itmX.SubItems(col) = "" & guest.InDate: col = col + 1
    itmX.SubItems(col) = "" & guest.OutDate: col = col + 1
    itmX.SubItems(col) = "" & nMaxTime: col = col + 1 '�Ҵ�ð�
    itmX.SubItems(col) = "" & guest.ParkTime: col = col + 1 '�����ð�
    
    nRemainTime = nMaxTime - guest.ParkTime
    If (nRemainTime > 0) Then
        itmX.SubItems(col) = "" & nRemainTime: col = col + 1  '�ܿ��ð�
        itmX.SubItems(col) = "" & 0: col = col + 1 '�ʰ��ð�
        
    Else
        itmX.SubItems(col) = "" & 0: col = col + 1 '�ܿ��ð�
        itmX.SubItems(col) = "" & Abs(nRemainTime): col = col + 1 '�ʰ��ð�
    End If

End Sub

Private Sub ClearGuestInfo(guest As stGuest)
    
    guest.CarGubun = "" '�湮����
    guest.Pass_YN = ""
    
    guest.InCarNo = ""
    guest.GuestName = ""
    guest.Dong = ""
    guest.Ho = ""
    guest.Tel = ""
    guest.object = ""
    guest.InGateNo = ""
    guest.InDate = ""
    guest.InImagePath = ""
    guest.RegDate = ""
    guest.ParkTime = ""
    
    guest.OutCarNo = ""
    guest.OutGateNo = ""
    guest.OutDate = ""
    guest.OutImagePath = ""
End Sub


'����Ű �Է½� �� ����
'���Ӽ� keypreview = true ����
Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        Call SSCommand6_Click
        KeyAscii = 0
        Sendkeys "{TAB}"
    End If
End Sub

Private Sub SSCommand3_Click()
    Me.MousePointer = 0
    FrmGuestRegCert.Show 1
    Call DataLogger("[HOST Button]    " & "�����湮���Խ��� ����")
End Sub


Private Sub SSCommand4_Click()
    FrmGuestRegCar.Show 1
End Sub

'��ȣ���� �����ð� ��ȸ
Private Sub SSCommand5_Click()
    'Call GuestRegParkTime_Daily
    Call GuestRegListView_DongHo
    Call SetDong
    Call SetHo
End Sub

'��������ȸ
Private Sub SSCommand6_Click()
    'Call GuestRegParkTime_Daily
    Call GuestRegListView_Car
    Call SetDong
    Call SetHo
End Sub

'���� �湮���� ���� �����ð� ���(������, ����, ȣ����)
Public Sub GuestRegParkTime_Daily()
'''    Dim sQry As String
'''    Dim sDong As String
'''    Dim sHo As String
'''    Dim sInCarno As String
'''    Dim sInDateTime As String
'''    Dim sOutCarno As String
'''    Dim sOutDateTime As String
'''    Dim nParkTIme As Long   '�����ð�
'''    Dim nParkCount As Long  '�����Ǽ�
'''    Dim nTotalParkTime As Long
'''    Dim nTotalParkCount As Long
'''    Dim nInSEQ As Long
'''    Dim nOutSEQ As Long
'''    Dim sName As String
'''    Dim sTel As String
'''    Dim bOUT_YN As Boolean '�ⱸ�ִ��� Ȯ��
'''
'''    Dim sStartDate As String
'''    Dim sEndDate As String
'''
'''
'''On Error GoTo Err_P
'''
'''    Call DataLogger("[�����湮����] ���������ð� ��� ����")
'''
'''    If ((LANE1_YN = "Y" And LANE1_Inout = "�ⱸ") Or (LANE2_YN = "Y" And LANE2_Inout = "�ⱸ") Or (LANE3_YN = "Y" And LANE3_Inout = "�ⱸ") Or (LANE4_YN = "Y" And LANE4_Inout = "�ⱸ") Or (LANE5_YN = "Y" And LANE5_Inout = "�ⱸ") Or (LANE6_YN = "Y" And LANE6_Inout = "�ⱸ")) Then
'''        bOUT_YN = True
'''    Else
'''        bOUT_YN = False
'''    End If
'''
'''    nParkCount = 0
'''    nTotalParkCount = 0
'''
'''    nParkTIme = 0
'''    nTotalParkTime = 0
'''
'''    '�ֱ� 2�� ������ ������� �湮���� ����ó������
'''    'sStartDate = DateAdd("m", (-2), Format(Now, "yyyy-mm-dd")) & " 00:00:00" '�ֱ� 2�� ��
'''    'sEndDate = Format(Now, "yyyy-mm-dd") & " 23:59:59"    '����
'''
'''    'sQry = "Select * from tb_guestReg_inout WHERE CAR_GUBUN='�湮����' AND (CALC IS NULL or CALC <> 'Y') "
'''    sQry = "Select * from tb_guestReg_inout WHERE CAR_NO <> '�νĽ���' AND (CALC IS NULL  OR  CALC <> 'Y') "
'''    sQry = sQry & " ORDER BY CAR_NO, PASS_DATE "
'''
'''    Set rs = New ADODB.Recordset
'''    rs.Open sQry, adoConn
'''    Do While Not (rs.EOF)
'''
'''        If (bOUT_YN = False) Then '�Ա��� ������
'''            If (rs!PASS_INOUT = "IN") Then  '�Ա�
'''                    nInSEQ = rs!SEQ
'''                    sDong = "" & rs!DRIVER_DEPT
'''                    sHo = "" & rs!DRIVER_CLASS
'''                    sInCarno = "" & rs!CAR_NO
'''                    sInDateTime = "" & Left(rs!PASS_DATE, 16)
'''                    sName = "" & rs!DRIVER_NAME
'''                    sTel = "" & rs!DRIVER_PHONE
'''
'''                    nParkTIme = 0
'''                    nParkCount = 1
'''                    nTotalParkTime = nTotalParkTime + nParkTIme
'''                    nTotalParkCount = nTotalParkCount + nParkCount
'''
'''                    adoConn.Execute "UPDATE tb_guestReg_inout SET CALC = 'Y' WHERE SEQ = '" & nInSEQ & "' LIMIT 1 "
'''                    adoConn.Execute "INSERT INTO tb_guestreg_daily (CAR_NO, DRIVER_DEPT, DRIVER_CLASS, IN_TIME, OUT_TIME, PARKTIME, PARKCOUNT, DRIVER_NAME, DRIVER_PHONE, REG_DATE) VALUES ('" & sInCarno & "','" & sDong & "','" & sHo & "', '" & sInDateTime & "', '" & sOutDateTime & "', " & nParkTIme & ", " & nParkCount & ", '" & sName & "', '" & sTel & "', '" & Left(rs!PASS_DATE, 19) & "')"
'''            End If
'''
'''
'''        Else '�Ա�/�ⱸ������
'''
'''            If (rs!PASS_INOUT = "IN") Then  '�Ա�
'''                nInSEQ = rs!SEQ
'''                sDong = "" & rs!DRIVER_DEPT
'''                sHo = "" & rs!DRIVER_CLASS
'''                sInCarno = "" & rs!CAR_NO
'''                sInDateTime = "" & Left(rs!PASS_DATE, 16)
'''                sName = "" & rs!DRIVER_NAME
'''                sTel = "" & rs!DRIVER_PHONE
'''
'''            ElseIf (rs!PASS_INOUT = "OUT") Then  '�ⱸ(��Ȯ�� ������ �ִ� ��츸 �����ð�/�Ǽ� �����)
'''                nOutSEQ = rs!SEQ
'''
'''                sDong = "" & rs!DRIVER_DEPT
'''                sHo = "" & rs!DRIVER_CLASS
'''
'''                sOutCarno = "" & rs!CAR_NO
'''                sOutDateTime = "" & Left(rs!PASS_DATE, 16)
'''
'''                If (sInCarno = sOutCarno) Then
'''                    nParkTIme = DateDiff("n", Left(sInDateTime, 19), Left(sOutDateTime, 19))
'''                    nTotalParkTime = nTotalParkTime + nParkTIme
'''
'''                    nParkCount = 1
'''                    nTotalParkCount = nTotalParkCount + 1
'''
'''                    'nMaxParkTime = GetMaxParkTime(sDong, sHo) '�����Ҵ�ð�
'''                    adoConn.Execute "INSERT INTO tb_guestreg_daily (CAR_NO, DRIVER_DEPT, DRIVER_CLASS, IN_TIME, OUT_TIME, PARKTIME, PARKCOUNT, DRIVER_NAME, DRIVER_PHONE, REG_DATE) VALUES ('" & sInCarno & "','" & sDong & "','" & sHo & "', '" & sInDateTime & "', '" & sOutDateTime & "', " & nParkTIme & ", " & nParkCount & ", '" & sName & "', '" & sTel & "', '" & Left(rs!PASS_DATE, 19) & "')"
'''                    adoConn.Execute "UPDATE tb_guestReg_inout SET CALC = 'Y' WHERE SEQ = '" & nInSEQ & "' LIMIT 1 "
'''                    adoConn.Execute "UPDATE tb_guestReg_inout SET CALC = 'Y' WHERE SEQ = '" & nOutSEQ & "' LIMIT 1 "
'''
'''                End If
'''
'''                nInSEQ = 0
'''                nOutSEQ = 0
'''                sInCarno = ""
'''                sInDateTime = ""
'''                sOutCarno = ""
'''                sOutDateTime = ""
'''                nParkTIme = 0
'''                nTotalParkTime = 0
'''                nParkCount = 0
'''            End If
'''        End If
'''
'''        rs.MoveNext
'''    Loop
'''
'''    Set rs = Nothing
'''
'''
'''
'''    Call DataLogger("[�����湮����] �����ð� ��� �Ϸ� �� ����")
'''
'''    Exit Sub
'''
'''Err_P:
'''    Set rs = Nothing
'''
'''    Call DataLogger("[�����湮����] �����ð� ��� ����:" & Err.Description)
End Sub


Private Sub GuestRegListView_Car()
    Dim Column_to_size As Integer
    Dim rs As Recordset
    Dim itmX As ListItem
    Dim INDEX_NO As Long
    Dim i As Long
    Dim bOUT_YN  As Boolean
    Dim nMaxTime As Long
    Dim nMaxParkTime As Long
    Dim nRemainTime As Long
    Dim sQry As String
    Dim sDong As String
    Dim sHo As String
    Dim sOldDong As String
    Dim sOldHo As String
    Dim nHomeParkTime As Long
    

    nHomeParkTime = 0
    
    If ((LANE1_YN = "Y" And LANE1_Inout = "�ⱸ") Or (LANE2_YN = "Y" And LANE2_Inout = "�ⱸ") Or (LANE3_YN = "Y" And LANE3_Inout = "�ⱸ") Or (LANE4_YN = "Y" And LANE4_Inout = "�ⱸ") Or (LANE5_YN = "Y" And LANE5_Inout = "�ⱸ") Or (LANE6_YN = "Y" And LANE6_Inout = "�ⱸ")) Then
        bOUT_YN = True
    Else
        bOUT_YN = False
    End If
    
    
    sQry = "SELECT * FROM tb_guestreg_daily WHERE " & " (REG_DATE >='" & Format(DTPicker1, "yyyy-mm-dd") & " 00:00:00.000') AND (REG_DATE <= '" & Format(DTPicker2, "yyyy-mm-dd") & " 23:59:59.999')"
    sDong = cmb_GuestDong.text
    sHo = cmb_GuestHo.text
    
    If (sDong <> "��ü") Then
        If (sHo <> "��ü") Then
            sQry = sQry & " AND DRIVER_DEPT = '" & sDong & "' AND DRIVER_CLASS = '" & sHo & "' "
        Else
            sQry = sQry & " AND DRIVER_DEPT = '" & sDong & "' "
        End If
    Else
        If (sHo <> "��ü") Then
            sQry = sQry & " AND DRIVER_CLASS = '" & sHo & "' "
        End If
    End If
    
    If (txt_Carno.text <> "") Then
        sQry = sQry & " AND (CAR_NO LIKE '%" & txt_Carno.text & "%')"
    End If
    
    sQry = sQry & " ORDER BY DRIVER_DEPT, DRIVER_CLASS "
    
'On Error GoTo Err_p

    Call ListViewExtended(ListView1)
    ListView1.View = lvwReport
    ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , " No "
    
    If (Glo_User_Type = "����1/����2") Then
        ListView1.ColumnHeaders.Add , , " �Ҽ�          "
        ListView1.ColumnHeaders.Add , , " ����          "
    Else
        ListView1.ColumnHeaders.Add , , " ��            "
        ListView1.ColumnHeaders.Add , , " ȣ��          "
    End If
   
    ListView1.ColumnHeaders.Add , , " ������ȣ           "
    ListView1.ColumnHeaders.Add , , " �湮��        "
    ListView1.ColumnHeaders.Add , , " ����ó                 "
    ListView1.ColumnHeaders.Add , , " �����ð�                "
    ListView1.ColumnHeaders.Add , , " �����ð�                "
    ListView1.ColumnHeaders.Add , , " �����ð�(��)"
    ListView1.ColumnHeaders.Add , , " "

    For Column_to_size = 0 To ListView1.ColumnHeaders.Count - 1
         SendMessage ListView1.hwnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
    Next

    INDEX_NO = 1
    Set rs = New ADODB.Recordset
    rs.Open sQry, adoConn
    Do While Not (rs.EOF)
        
        
        If (rs!DRIVER_DEPT = sOldDong And rs!DRIVER_CLASS = sOldHo) Then
            nHomeParkTime = nHomeParkTime + rs!ParkTime '���� ��ȣ���� �����ð� ����
        Else
            nHomeParkTime = 0 '��ȣ���� �ٸ��� �����ð� �ʱ�ȭ
            sOldDong = rs!DRIVER_DEPT
            sOldHo = rs!DRIVER_CLASS
        End If
        
        
        Set itmX = ListView1.ListItems.Add(, , "" & INDEX_NO)
        
        i = 1
        itmX.SubItems(i) = "" & rs!DRIVER_DEPT: i = i + 1
        itmX.SubItems(i) = "" & rs!DRIVER_CLASS: i = i + 1
        itmX.SubItems(i) = "" & rs!CAR_NO: i = i + 1
        itmX.SubItems(i) = "" & rs!DRIVER_NAME: i = i + 1
        itmX.SubItems(i) = "" & rs!DRIVER_PHONE: i = i + 1
        itmX.SubItems(i) = "" & rs!IN_TIME: i = i + 1
        
        If (bOUT_YN = True) Then '�ⱸ���� ��
            itmX.SubItems(i) = "" & rs!OUT_TIME: i = i + 1
            itmX.SubItems(i) = "" & rs!ParkTime: i = i + 1
            
        Else
            itmX.SubItems(i) = "" '�����ð�
            itmX.SubItems(i) = "" '�����ð�
        End If
        
        

        INDEX_NO = INDEX_NO + 1

        rs.MoveNext
    Loop
    Set rs = Nothing
    
    LblRecordCount.Caption = INDEX_NO - 1
    
Exit Sub
End Sub

'�����ð� ��ȸ
Private Sub GuestRegListView_DongHo()
    Dim Column_to_size As Integer
    Dim rs As Recordset
    Dim itmX As ListItem
    Dim INDEX_NO As Long
    Dim i As Long
    Dim bOUT_YN  As Boolean
    Dim nMaxTime As Long
    Dim nMaxParkTime As Long
    Dim nRemainTime As Long
    Dim sQry As String
    Dim sDong As String
    Dim sHo As String
    Dim sOldDong As String
    Dim sOldHo As String
    Dim nHomeParkTime As Long
    

    nHomeParkTime = 0
    
    If ((LANE1_YN = "Y" And LANE1_Inout = "�ⱸ") Or (LANE2_YN = "Y" And LANE2_Inout = "�ⱸ") Or (LANE3_YN = "Y" And LANE3_Inout = "�ⱸ") Or (LANE4_YN = "Y" And LANE4_Inout = "�ⱸ") Or (LANE5_YN = "Y" And LANE5_Inout = "�ⱸ") Or (LANE6_YN = "Y" And LANE6_Inout = "�ⱸ")) Then
        bOUT_YN = True
    Else
        bOUT_YN = False
    End If
    
    
    'sQry = "SELECT *, SUM(PARKTIME) as SUMPTIME, SUM(PARKCOUNT) as SUMPCOUNT FROM tb_guestreg_daily WHERE " & " (REG_DATE >='" & Format(DTPicker1, "yyyy-mm-dd") & " 00:00:00.000') AND (REG_DATE <= '" & Format(DTPicker2, "yyyy-mm-dd") & " 23:59:59.999')"
    sQry = "SELECT *, SUM(PARKTIME) as SUMPTIME, COUNT(*) AS SUMPCOUNT FROM tb_guestreg_daily WHERE " & " (REG_DATE >='" & Format(DTPicker1, "yyyy-mm-dd") & " 00:00:00.000') AND (REG_DATE <= '" & Format(DTPicker2, "yyyy-mm-dd") & " 23:59:59.999')"
    sDong = cmb_GuestDong.text
    sHo = cmb_GuestHo.text
        
    If (sDong <> "��ü") Then
        If (sHo <> "��ü") Then
            sQry = sQry & " AND DRIVER_DEPT = '" & sDong & "' AND DRIVER_CLASS = '" & sHo & "' "
        Else
            sQry = sQry & " AND DRIVER_DEPT = '" & sDong & "' "
        End If
    Else
        If (sHo <> "��ü") Then
            sQry = sQry & " AND DRIVER_CLASS = '" & sHo & "' "
        End If
    End If
    
    If (txt_Carno.text <> "") Then
        txt_Carno = ""
    End If
    
    'sQry = sQry & " ORDER BY DRIVER_DEPT, DRIVER_CLASS "
    sQry = sQry & " GROUP BY DRIVER_DEPT, DRIVER_CLASS "
    
'On Error GoTo Err_p

    Call ListViewExtended(ListView1)
    ListView1.View = lvwReport
    ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , " No "
    
    If (Glo_User_Type = "����1/����2") Then
        ListView1.ColumnHeaders.Add , , " �Ҽ�          "
        ListView1.ColumnHeaders.Add , , " ����          "
    Else
        ListView1.ColumnHeaders.Add , , " ��            "
        ListView1.ColumnHeaders.Add , , " ȣ��          "
    End If
    
    
    ListView1.ColumnHeaders.Add , , " �� �Ҵ�ð�(��)"
    ListView1.ColumnHeaders.Add , , " �����ð�(��)" '��ȣ���� �����ð� �հ�
    ListView1.ColumnHeaders.Add , , " �����ð�(��)"
    ListView1.ColumnHeaders.Add , , " �����Ǽ�"
    ListView1.ColumnHeaders.Add , , " ó���Ͻ�                     "
    
    ListView1.ColumnHeaders.Add , , " "

    For Column_to_size = 0 To ListView1.ColumnHeaders.Count - 1
         SendMessage ListView1.hwnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
    Next

    INDEX_NO = 1
    Set rs = New ADODB.Recordset
    rs.Open sQry, adoConn
    Do While Not (rs.EOF)
        
        
        If (rs!DRIVER_DEPT = sOldDong And rs!DRIVER_CLASS = sOldHo) Then
            nHomeParkTime = nHomeParkTime + rs!ParkTime '���� ��ȣ���� �����ð� ����
        Else
            nHomeParkTime = 0 '��ȣ���� �ٸ��� �����ð� �ʱ�ȭ
            sOldDong = rs!DRIVER_DEPT
            sOldHo = rs!DRIVER_CLASS
        End If

        Set itmX = ListView1.ListItems.Add(, , "" & INDEX_NO)
        
        i = 1
        itmX.SubItems(i) = "" & rs!DRIVER_DEPT: i = i + 1
        itmX.SubItems(i) = "" & rs!DRIVER_CLASS: i = i + 1
        
        nMaxParkTime = GetMaxParkTime("" & rs!DRIVER_DEPT, "" & rs!DRIVER_CLASS) '�� �Ҵ�ð�(��)
        itmX.SubItems(i) = "" & nMaxParkTime: i = i + 1
        
        If (bOUT_YN = True) Then '�ⱸ���� ��

            itmX.SubItems(i) = "" & rs!SUMPTIME: i = i + 1 '�����ð� �հ�
            
            nRemainTime = nMaxParkTime - rs!SUMPTIME 'nAccParkTime:���� ��ȣ�� �϶� �����ð�
            itmX.SubItems(i) = "" & nRemainTime: i = i + 1  '�����ð�
            
        Else
            itmX.SubItems(i) = "": i = i + 1 '�����ð�
            itmX.SubItems(i) = "": i = i + 1 '�����ð�

        End If
        
        itmX.SubItems(i) = "" & rs!SUMPCOUNT: i = i + 1  '�����Ǽ�
        itmX.SubItems(i) = "" & rs!REG_DATE: i = i + 1 'ó���Ͻ�
        

        INDEX_NO = INDEX_NO + 1


        rs.MoveNext
    Loop
    Set rs = Nothing
    
    LblRecordCount.Caption = INDEX_NO - 1
    
Exit Sub
End Sub

'�湮���� ���� ��ȸ/����
Private Sub SSCommand7_Click()
    Me.MousePointer = 0
    FrmGuestRegLimit.Show 1
    Call DataLogger("[HOST Button]    " & "�湮���� ����")
End Sub

'�̽��� ������ �ִ��� Ȯ��
Private Sub Timer_CheckSignup_Timer()
    Dim bQryResult As Boolean
    Dim rs As ADODB.Recordset
    Dim sQry As String
    
    
    Set rs = New ADODB.Recordset
    bQryResult = DataBaseQuery(rs, adoConn, "SELECT count(USE_YN) as NotUse_Cnt from tb_guestreg_admin WHERE USE_YN <>'Y';", False)
    If (bQryResult = False) Then
        Exit Sub
    End If
    
    If (rs!NotUse_Cnt > 0) Then
        SSCommand3.ForeColor = &HFFFF& '���
    Else
        SSCommand3.ForeColor = &HFFFFFF '���
    End If

    Set rs = Nothing
    
End Sub
