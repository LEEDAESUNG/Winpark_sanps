VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMctl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmWebdc 
   BackColor       =   &H00404040&
   BorderStyle     =   1  '���� ����
   Caption         =   "ParkingManager��"
   ClientHeight    =   10740
   ClientLeft      =   5160
   ClientTop       =   1725
   ClientWidth     =   17655
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   10740
   ScaleWidth      =   17655
   Begin VB.CommandButton cmd_AutoFreeCharge 
      Appearance      =   0  '���
      BackColor       =   &H8000000A&
      Caption         =   "       ��������Ʈ      �ڵ����� ����"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   13245
      Style           =   1  '�׷���
      TabIndex        =   22
      ToolTipText     =   "�ſ� 1�� ��������Ʈ�� �ڵ������մϴ�"
      Top             =   2295
      Width           =   1830
   End
   Begin VB.ComboBox cmb_StoreCharge 
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
      Left            =   13245
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   16
      Top             =   1770
      Width           =   3720
   End
   Begin VB.ComboBox cmb_GroupCharge 
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
      Left            =   10425
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   15
      Top             =   1770
      Width           =   1830
   End
   Begin VB.CommandButton cmd_FreeCharge 
      Appearance      =   0  '���
      BackColor       =   &H8000000A&
      Caption         =   "��������Ʈ ����"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   15150
      Style           =   1  '�׷���
      TabIndex        =   14
      Top             =   2295
      Width           =   1830
   End
   Begin VB.TextBox txt_FreePoint 
      BeginProperty Font 
         Name            =   "�������"
         Size            =   14.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   10425
      TabIndex        =   13
      Text            =   "txt_FreePoint"
      Top             =   2310
      Width           =   1830
   End
   Begin VB.TextBox txt_PaidPointMoney 
      BeginProperty Font 
         Name            =   "�������"
         Size            =   14.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   13245
      TabIndex        =   12
      Text            =   "txt_PaidPoint_ChargeMoney"
      Top             =   2985
      Width           =   1830
   End
   Begin VB.CommandButton cmd_PaidCharge 
      Appearance      =   0  '���
      BackColor       =   &H8000000A&
      Caption         =   "��������Ʈ ����"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   15150
      MaskColor       =   &H00FFFFFF&
      Style           =   1  '�׷���
      TabIndex        =   11
      Top             =   2985
      Width           =   1830
   End
   Begin VB.TextBox txt_PaidPoint 
      BeginProperty Font 
         Name            =   "�������"
         Size            =   14.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   10425
      TabIndex        =   10
      Text            =   "txt_PaidPoint"
      Top             =   2985
      Width           =   1830
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   5265
      Left            =   30
      TabIndex        =   0
      Top             =   4440
      Width           =   17595
      _ExtentX        =   31036
      _ExtentY        =   9287
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   345
      Left            =   13050
      TabIndex        =   5
      Top             =   1125
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
      Format          =   253296640
      CurrentDate     =   36927
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   345
      Left            =   15150
      TabIndex        =   6
      Top             =   1125
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
      Format          =   253296640
      CurrentDate     =   36927
   End
   Begin Threed.SSCommand SSCommand3 
      Height          =   615
      Left            =   15135
      TabIndex        =   7
      ToolTipText     =   "����Ʈ ���� ���� ��ȸ"
      Top             =   3705
      Width           =   1830
      _Version        =   65536
      _ExtentX        =   3228
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "��������"
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�������"
         Size            =   14.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9705
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Threed.SSCommand SSCommand2 
      Height          =   570
      Left            =   15720
      TabIndex        =   8
      Top             =   225
      Width           =   1260
      _Version        =   65536
      _ExtentX        =   2222
      _ExtentY        =   1005
      _StockProps     =   78
      Caption         =   "�� ��"
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
      Picture         =   "FrmWebDC.frx":0000
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   570
      Left            =   14400
      TabIndex        =   9
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
      Picture         =   "FrmWebDC.frx":0351
   End
   Begin Threed.SSCommand SSCommand4 
      Height          =   615
      Left            =   13260
      TabIndex        =   27
      ToolTipText     =   "��ü�� ������ ���� ��ȸ"
      Top             =   3705
      Width           =   1830
      _Version        =   65536
      _ExtentX        =   3228
      _ExtentY        =   1085
      _StockProps     =   78
      Caption         =   "������ ��ȸ"
      ForeColor       =   4210752
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�������"
         Size            =   14.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      X1              =   16980
      X2              =   9195
      Y1              =   2910
      Y2              =   2910
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "��������Ʈ �ڵ����� : "
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
      Index           =   3
      Left            =   13725
      TabIndex        =   28
      Top             =   10065
      Visible         =   0   'False
      Width           =   2745
   End
   Begin VB.Label lbl_COUNT 
      BackStyle       =   0  '����
      Caption         =   "��ȸ�Ǽ� :"
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
      Left            =   345
      TabIndex        =   26
      Top             =   10065
      Width           =   1425
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�ݾ� : "
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
      Index           =   2
      Left            =   10170
      TabIndex        =   25
      Top             =   10065
      Width           =   780
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "��������Ʈ : "
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
      Index           =   1
      Left            =   6735
      TabIndex        =   24
      Top             =   10065
      Width           =   1590
   End
   Begin VB.Label lbl_option 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "��������Ʈ : "
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
      Left            =   3285
      TabIndex        =   23
      Top             =   10065
      Width           =   1590
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "��������Ʈ"
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
      Left            =   9240
      TabIndex        =   21
      Top             =   3105
      Width           =   1125
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "��������Ʈ"
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
      Left            =   9225
      TabIndex        =   20
      Top             =   2430
      Width           =   1125
   End
   Begin VB.Label Label3 
      BackColor       =   &H00404040&
      Caption         =   "�ݾ�"
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
      Height          =   315
      Left            =   12675
      TabIndex        =   19
      Top             =   3120
      Width           =   525
   End
   Begin VB.Label Label2 
      Alignment       =   1  '������ ����
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "��ü��"
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
      Index           =   1
      Left            =   12510
      TabIndex        =   18
      Top             =   1800
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�׷�"
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
      Left            =   9900
      TabIndex        =   17
      Top             =   1800
      Width           =   450
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '����
      DataField       =   "imgpath1"
      DataSource      =   "Adodc1"
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   255
      TabIndex        =   4
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
      TabIndex        =   3
      Top             =   1155
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
      TabIndex        =   2
      Top             =   1155
      Width           =   150
   End
   Begin VB.Label lbl_APS 
      BackStyle       =   0  '����
      Caption         =   "������ ��ȸ �� ����"
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
      Index           =   0
      Left            =   210
      TabIndex        =   1
      Top             =   480
      Width           =   4815
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
End
Attribute VB_Name = "FrmWebdc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sButton_Search As String

'�޺��ڽ�:�׷켱�ý� ���弱��
Private Sub cmb_GroupCharge_Click()
    Call SetCombo_ChargeStore(cmb_GroupCharge.text)
End Sub

'�ش� �ط��� ���帮��Ʈ ���
Private Sub SetCombo_ChargeStore(sGroupName As String)
    Dim rs As Recordset
    Dim rs2 As Recordset
    Dim sQry As String
    Dim sSEQ As String
    
On Error GoTo Err_P
    
    cmb_StoreCharge.Clear
    cmb_StoreCharge.AddItem "��ü"
    
    If (sGroupName = "��ü") Then
        sQry = "SELECT * FROM tb_id WHERE GUBUN != '�Ѱ�������' AND GUBUN != '������' AND GUBUN != '���' " '��� �׷�(������ �ش�)
    Else
        sQry = "SELECT * FROM tb_id WHERE GUBUN = '" & sGroupName & "' " 'Ư�� �׷�
    End If

    Set rs = New ADODB.Recordset
    rs.Open sQry, adoConn
    Do While Not (rs.EOF)

        sSEQ = "" & rs!SEQ
        
        Set rs2 = New ADODB.Recordset
        sQry = "SELECT * FROM tb_partner WHERE SEQ = '" & sSEQ & "'"
        rs2.Open sQry, adoConn
        If Not (rs2.EOF) Then
            cmb_StoreCharge.AddItem rs2!SEQ & "." & rs2!ID & "(" & rs2!PNAME & ")"
        End If
        Set rs2 = Nothing
        
        rs.MoveNext
    Loop

    Set rs2 = Nothing
    Set rs = Nothing
    
    cmb_StoreCharge.ListIndex = 0
    
    Exit Sub
    
Err_P:
    Set rs2 = Nothing
    Set rs = Nothing

    Call DataLogger("[FrmWebdc SetCombo_PartnerName]    " & sGroupName & ":�׷�� ���� �� ����� ��¿���. �ٽ� �õ����ּ���(E00006)" & " " & Err.Description)
End Sub

'�ڵ� �������� ����
Private Sub cmd_AutoFreeCharge_Click()
    Dim rs As Recordset
    Dim rs2 As Recordset
    Dim sQry As String
    Dim sQry2 As String
    Dim bQryResult As Boolean
    Dim nAutoFreePoint As Integer
    Dim sSEQ, sID, sPSEQ As String
    Dim sLog As String
    Dim sStrLine() As String
    Dim sRegDate As String
    
    

On Error GoTo Err_P
    
    If (CheckFreeChargeValue = False) Then
        Msg_Box.Label2.Caption = "�Է¿���"
        Msg_Box.Label1.Caption = "���ڸ� �Է��ϼ���."
        Msg_Box.Show 1
        Exit Sub
    End If


    MBox.Label2.Caption = "������"
    MBox.Label3.Caption = cmb_GroupCharge.text
    MBox.Label1.Caption = "�ڵ��������� �����Ͻðڽ��ϱ�?" & vbCrLf & vbCrLf & "�Ŵ� 1�� �ڵ����� �������� �˴ϴ�." & vbCrLf & vbCrLf & "�����Ͻðڽ��ϱ�?"
    MBox.Show 1
    If (Glo_MsgRet = False) Then
        Set rs = Nothing
        Exit Sub
    End If

    Call DataLogger("[FrmWebdc AutoFreeCharge]    " & "�ڵ��������� ���� ��ư")
    sQry = "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('', 'HOST','�ڵ��������� ���� ��ư','" & Glo_Login_ID & "'," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
    adoConn.Execute sQry
    
    sStrLine = Split(cmb_StoreCharge.text, ".")
    sID = sStrLine(0)
    
    If (cmb_GroupCharge.text = "��ü") Then
        If (cmb_StoreCharge.text = "��ü") Then
            sQry = "SELECT * FROM tb_id WHERE GUBUN != '�Ѱ�������' AND GUBUN != '������' AND GUBUN != '���' " '��� ��Ʈ��
        Else
            sQry = "SELECT * FROM tb_id WHERE SEQ = '" & sID & "' " 'Ư�� ID
        End If
    Else
        If (cmb_StoreCharge.text = "��ü") Then
            sQry = "SELECT * FROM tb_id WHERE GUBUN = '" & cmb_GroupCharge.text & "' " 'Ư�� �׷�
        Else
            sQry = "SELECT * FROM tb_id WHERE SEQ = '" & sID & "' " 'Ư�� ID
        End If
    End If
    
    sRegDate = Format(Now, "yyyy-mm-dd hh:nn:ss")
    
    Set rs = New ADODB.Recordset
    rs.Open sQry, adoConn
    Do While Not (rs.EOF)
        
            sSEQ = "" & rs!SEQ
            sID = "" & rs!ID
            
            sQry = "SELECT * FROM tb_partner WHERE SEQ = '" & sSEQ & "'"
            Set rs2 = New ADODB.Recordset
            rs2.Open sQry, adoConn
            If Not (rs2.EOF) Then

                nAutoFreePoint = txt_FreePoint.text
                If (nAutoFreePoint < 0) Then
                    nAutoFreePoint = 0
                End If
                
                sLog = "[������ �ڵ���������]" & sSEQ & "." & sID & "(" & "):" & nAutoFreePoint & "(��)"
                
                If (cmb_StoreCharge.text = "��ü") Then
    
                    sQry = "UPDATE  tb_partner  SET  FREE_AUTOPOINT = " & nAutoFreePoint & ", FREE_AUTOPOINT_LASTDATE = '" & sRegDate & "' WHERE SEQ = '" & sSEQ & "' "
                    adoConn.Execute sQry
                    
                    sQry = "INSERT INTO tb_partner_log (PCODE, FREE_POINT, INFO, CHARGE_ACCOUNT, REG_DATE) values ('" & sSEQ & "', " & nAutoFreePoint & ", '" & sLog & "', '" & Glo_Login_ID & "', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "' )"
                    adoConn.Execute sQry
                    
                    sQry = "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('" & sSEQ & "', 'HOST','" & sLog & "','" & Glo_Login_ID & "'," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
                    adoConn.Execute sQry
                    
                    'List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & sLog, 0
                    Call DataLogger("[FrmWebdc FreeCharge]    " & sLog)
                    
                Else
                    sStrLine() = Split(cmb_StoreCharge.text, ".")
                    sPSEQ = sStrLine(0)
                    
                    If (sSEQ = sPSEQ) Then
                        sQry = "UPDATE  tb_partner  SET  FREE_AUTOPOINT = " & nAutoFreePoint & ", FREE_AUTOPOINT_LASTDATE = '" & sRegDate & "' WHERE SEQ = '" & sPSEQ & "' "
                        adoConn.Execute sQry
                        
                        sQry = "INSERT INTO tb_partner_log (PCODE, FREE_POINT, INFO, CHARGE_ACCOUNT, REG_DATE) values ('" & sSEQ & "', " & nAutoFreePoint & ", '" & sLog & "', '" & Glo_Login_ID & "', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "' )"
                        adoConn.Execute sQry
                        
                        sQry = "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('" & sSEQ & "', 'HOST','" & sLog & "','" & Glo_Login_ID & "'," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
                        adoConn.Execute sQry
                        
                        Call DataLogger("[FrmWebdc FreeCharge]    " & sLog)
                        
                        Exit Do
                    End If
                End If
                
            Else
                Set rs2 = Nothing
            End If
            

        rs.MoveNext
    Loop

    Set rs2 = Nothing
    Set rs = Nothing
    
    Call DataLogger("[FrmWebdc AutoFreeCharge]    " & "��������Ʈ �ڵ����� ���� �Ϸ��߽��ϴ�")
    Msg_Box.Label2.Caption = "������"
    Msg_Box.Label1.Caption = "��������Ʈ �ڵ����� ���� �Ϸ��߽��ϴ�"
    Msg_Box.Show 1
    
    Exit Sub
    
Err_P:
    Set rs = Nothing
    
    'List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_id & ":�����Ϳ����߻�. �ٽ� �õ����ּ���(E00004)" & " " & Err.Description, 0
    'Call DataLogger("[FrmWebdc FreeCharge]    " & txt_id & ":�����Ϳ����߻�. �ٽ� �õ����ּ���(E00004)" & " " & Err.Description)
    Call DataLogger("[FrmWebdc AutoFreeCharge]    " & cmb_GroupCharge.text & "_" & cmb_StoreCharge & "_" & txt_FreePoint & ":�ڵ��������� ��������. �ٽ� �õ����ּ���(E00007)" & " " & Err.Description)
End Sub

Private Sub cmd_FreeCharge_Click()
    
    Dim rs As Recordset
    Dim rs2 As Recordset
    Dim sQry As String
    Dim sQry2 As String
    Dim bQryResult As Boolean
    Dim nFreePoint, nAddFreePoint, nSumFreePoint As Integer
    Dim sSEQ, sID, sPSEQ As String
    Dim sLog As String
    Dim sStrLine() As String
    
On Error GoTo Err_P
            
    If (CheckFreeChargeValue = False) Then
        Msg_Box.Label2.Caption = "�Է¿���"
        Msg_Box.Label1.Caption = "���ڸ� �Է��ϼ���."
        Msg_Box.Show 1
        Exit Sub
    End If
    
    MBox.Label2.Caption = "������"
    MBox.Label3.Caption = cmb_GroupCharge.text
    MBox.Label1.Caption = "�������� �����Ͻðڽ��ϱ�?"
    MBox.Show 1
    If (Glo_MsgRet = False) Then
        Exit Sub
    End If
    
    Call DataLogger("[FrmWebdc]    " & "�������� ��ư")
    sQry = "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('', 'HOST','�������� ��ư','" & Glo_Login_ID & "'," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
    adoConn.Execute sQry
    
    
    If (cmb_GroupCharge.text = "��ü") Then
        sQry = "SELECT * FROM tb_id WHERE GUBUN != '�Ѱ�������' AND GUBUN != '������' AND GUBUN != '���'" '��� �׷�(������ �ش�)
    Else
        sQry = "SELECT * FROM tb_id WHERE GUBUN = '" & cmb_GroupCharge.text & "' " 'Ư�� �׷�
    End If

    
    Set rs = New ADODB.Recordset
    rs.Open sQry, adoConn
    Do While Not (rs.EOF)
        
            sSEQ = "" & rs!SEQ
            sID = "" & rs!ID
            
            sQry = "SELECT * FROM tb_partner WHERE SEQ = '" & sSEQ & "'"
            Set rs2 = New ADODB.Recordset
            rs2.Open sQry, adoConn
            If Not (rs2.EOF) Then

                nFreePoint = rs2!FREE_POINT
                nAddFreePoint = CInt(txt_FreePoint.text)
                nSumFreePoint = nFreePoint + nAddFreePoint
                If (nSumFreePoint < 0) Then
                    nSumFreePoint = 0
                End If
                
                sLog = "[������ ��������]" & sSEQ & "." & sID & "(" & "):" & nAddFreePoint & "(��)"
                
                If (cmb_StoreCharge.text = "��ü") Then
    
                    sQry = "UPDATE  tb_partner  SET  FREE_POINT = " & nSumFreePoint & " WHERE SEQ = '" & sSEQ & "' "
                    adoConn.Execute sQry
                    
                    sQry = "INSERT INTO tb_partner_log (PCODE, FREE_POINT, INFO, CHARGE_ACCOUNT, REG_DATE) values ('" & sSEQ & "', " & nAddFreePoint & ", '" & sLog & "', '" & Glo_Login_ID & "', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "' )"
                    adoConn.Execute sQry
                    
                    sQry = "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('" & sSEQ & "', 'HOST','" & sLog & "','" & Glo_Login_ID & "'," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
                    adoConn.Execute sQry
                    
                    'List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & sLog, 0
                    Call DataLogger("[FrmWebdc FreeCharge]    " & sLog)
                    
                Else
                    sStrLine() = Split(cmb_StoreCharge.text, ".")
                    sPSEQ = sStrLine(0)
                    
                    If (sSEQ = sPSEQ) Then
                        sQry = "UPDATE  tb_partner  SET  FREE_POINT = " & nSumFreePoint & " WHERE SEQ = '" & sPSEQ & "' "
                        adoConn.Execute sQry
                        
                        sQry = "INSERT INTO tb_partner_log (PCODE, FREE_POINT, INFO, CHARGE_ACCOUNT, REG_DATE) values ('" & sSEQ & "', " & nAddFreePoint & ", '" & sLog & "', '" & Glo_Login_ID & "', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "' )"
                        adoConn.Execute sQry
                        
                        sQry = "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('" & sSEQ & "', 'HOST','" & sLog & "','" & Glo_Login_ID & "'," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
                        adoConn.Execute sQry
                        
                        Call DataLogger("[FrmWebdc FreeCharge]    " & sLog)
                        
                        Exit Do
                    End If
                End If
                
            Else
                Set rs2 = Nothing
            End If
            

        rs.MoveNext
    Loop

    Set rs2 = Nothing
    Set rs = Nothing
    
    Call DataLogger("[FrmWebdc FreeCharge]    " & "��������Ʈ ���� �Ϸ��߽��ϴ�")
    Msg_Box.Label2.Caption = "������"
    Msg_Box.Label1.Caption = "��������Ʈ ���� �Ϸ��߽��ϴ�"
    Msg_Box.Show 1

    
    Exit Sub
    
Err_P:
    Set rs2 = Nothing
    Set rs = Nothing
    
    'List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_id & ":�����Ϳ����߻�. �ٽ� �õ����ּ���(E00004)" & " " & Err.Description, 0
    'Call DataLogger("[FrmWebdc FreeCharge]    " & txt_id & ":�����Ϳ����߻�. �ٽ� �õ����ּ���(E00004)" & " " & Err.Description)
    Call DataLogger("[FrmWebdc FreeCharge]    " & cmb_GroupCharge.text & "_" & cmb_StoreCharge & "_" & txt_FreePoint & ":�������� ����. �ٽ� �õ����ּ���(E00006)" & " " & Err.Description)
End Sub


'�������� ������ üũ
Private Function CheckPaidChargeValue()
    Dim bCheck As Boolean
    
    bCheck = True
    
    txt_PaidPoint = Trim(txt_PaidPoint)
    txt_PaidPointMoney = Trim(txt_PaidPointMoney)
    
    If (txt_PaidPoint = "") Then txt_PaidPoint = "0"
    If (txt_PaidPointMoney = "") Then txt_PaidPointMoney = "0"
    
    If (IsNumeric(txt_PaidPoint) = False) Then
        txt_PaidPoint = "0"
        txt_PaidPoint.SetFocus
        bCheck = False
    End If
    If (IsNumeric(txt_PaidPointMoney) = False) Then
        txt_PaidPointMoney = "0"
        txt_PaidPointMoney.SetFocus
        bCheck = False
    End If
    
    'If (CInt(txt_PaidPoint) = 0) Then
    '    bCheck = False
    'End If
    
    CheckPaidChargeValue = bCheck
    
    Exit Function
Err_P:
    CheckPaidChargeValue = False
End Function

'�������� ������ üũ
Private Function CheckFreeChargeValue()
    Dim bCheck As Boolean
    
    bCheck = True
    
    txt_FreePoint = Trim(txt_FreePoint)
    
    If (txt_FreePoint = "") Then txt_FreePoint = "0"
    
    If (IsNumeric(txt_FreePoint) = False) Then
        txt_FreePoint = "0"
        txt_FreePoint.SetFocus
        bCheck = False
    End If
    
    'If (CInt(txt_FreePoint) < 0) Then
    '    bCheck = False
    'End If
    
    CheckFreeChargeValue = bCheck
    
    Exit Function
Err_P:
    CheckFreeChargeValue = False
End Function

Private Sub cmd_PaidCharge_Click()
    Dim rs As Recordset
    Dim rs2 As Recordset
    Dim sQry As String
    Dim bQryResult As Boolean
    Dim nPaidPoint, nAddPaidPoint, nSumPaidPoint As Integer
    Dim nPaidPoint_Money As Long
    Dim sSEQ, sID, sPSEQ As String
    Dim sLog As String
    Dim sStrLine() As String
    
On Error GoTo Err_P
    
    If (CheckPaidChargeValue = False) Then
        Msg_Box.Label2.Caption = "�Է¿���"
        Msg_Box.Label1.Caption = "���ڸ� �Է��ϼ���."
        Msg_Box.Show 1
        Exit Sub
    End If
    
    MBox.Label2.Caption = "������"
    MBox.Label3.Caption = cmb_GroupCharge.text
    MBox.Label1.Caption = "�������� �����Ͻðڽ��ϱ�?"
    MBox.Show 1
    If (Glo_MsgRet = False) Then
        Exit Sub
    End If
    
    Call DataLogger("[FrmWebdc]    " & "�������� ��ư")
    sQry = "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('', 'HOST','�������� ��ư','" & Glo_Login_ID & "'," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
    adoConn.Execute sQry
    
    
    If (cmb_GroupCharge.text = "��ü") Then
        sQry = "SELECT * FROM tb_id WHERE GUBUN != '�Ѱ�������' AND GUBUN != '������' AND GUBUN != '���'" '��� �׷�(������ �ش�)
    Else
        sQry = "SELECT * FROM tb_id WHERE GUBUN = '" & cmb_GroupCharge.text & "' " 'Ư�� �׷�
    End If
    
    
    Set rs = New ADODB.Recordset
    rs.Open sQry, adoConn
    Do While Not (rs.EOF)
            
            sSEQ = "" & rs!SEQ
            sID = "" & rs!ID
            
            
            sQry = "SELECT * FROM tb_partner WHERE SEQ = '" & sSEQ & "'"
            Set rs2 = New ADODB.Recordset
            rs2.Open sQry, adoConn
            If Not (rs2.EOF) Then
             
                nPaidPoint = rs2!PAID_POINT
                nAddPaidPoint = CInt(txt_PaidPoint.text)
                nSumPaidPoint = nPaidPoint + nAddPaidPoint
                nPaidPoint_Money = "" & txt_PaidPointMoney
                If (nSumPaidPoint < 0) Then
                    nSumPaidPoint = 0
                End If
                
                sLog = "[������ ��������]" & sID & ":" & nAddPaidPoint & "(��)"
                
                If (cmb_StoreCharge.text = "��ü") Then
                    sQry = "UPDATE  tb_partner  SET  PAID_POINT = " & nSumPaidPoint & " WHERE SEQ = '" & sSEQ & "' "
                    adoConn.Execute sQry
                    
                    sQry = "INSERT INTO tb_partner_log (PCODE, PAID_POINT, PAID_POINT_CHARGEMONEY, INFO, CHARGE_ACCOUNT, REG_DATE) values ('" & sSEQ & "', " & nAddPaidPoint & ", " & nPaidPoint_Money & ", '" & sLog & "', '" & Glo_Login_ID & "', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "' )"
                    adoConn.Execute sQry
                    
                    sQry = "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('" & sSEQ & "', 'HOST','" & sLog & "','" & Glo_Login_ID & "'," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
                    adoConn.Execute sQry
                    
                    'List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & sLog, 0
                    Call DataLogger("[FrmWebdc PaidCharge]    " & sLog)
                
                Else
                    sStrLine() = Split(cmb_StoreCharge.text, ".")
                    sPSEQ = sStrLine(0)
                    
                    If (sSEQ = sPSEQ) Then
                        sQry = "UPDATE  tb_partner  SET  PAID_POINT = " & nSumPaidPoint & " WHERE SEQ = '" & sPSEQ & "' "
                        adoConn.Execute sQry
                        
                        sQry = "INSERT INTO tb_partner_log (PCODE, PAID_POINT, PAID_POINT_CHARGEMONEY, INFO, CHARGE_ACCOUNT, REG_DATE) values ('" & sSEQ & "', " & nAddPaidPoint & ", " & nPaidPoint_Money & ", '" & sLog & "', '" & Glo_Login_ID & "', '" & Format(Now, "YYYY-MM-DD HH:NN:SS") & "' )"
                        adoConn.Execute sQry
                        
                        sQry = "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('" & sSEQ & "', 'HOST','" & sLog & "','" & Glo_Login_ID & "'," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
                        adoConn.Execute sQry

                        Call DataLogger("[FrmWebdc PaidCharge]    " & sLog)
                        Exit Do
                    End If
                End If
            Else
                Set rs2 = Nothing
            End If
            
            rs.MoveNext
    Loop

    Set rs2 = Nothing
    Set rs = Nothing
    
    Call DataLogger("[FrmWebdc PaidCharge]    " & "��������Ʈ ���� �Ϸ��߽��ϴ�")
    Msg_Box.Label2.Caption = "������"
    Msg_Box.Label1.Caption = "��������Ʈ ���� �Ϸ��߽��ϴ�"
    Msg_Box.Show 1
    
    Exit Sub
    
Err_P:
    Set rs2 = Nothing
    Set rs = Nothing
    
    'List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_id & ":�����Ϳ����߻�. �ٽ� �õ����ּ���(E00004)" & " " & Err.Description, 0
    'Call DataLogger("[FrmWebdc PaidCharge]    " & txt_id & ":�����Ϳ����߻�. �ٽ� �õ����ּ���(E00004)" & " " & Err.Description)
    Call DataLogger("[FrmWebdc PaidCharge]    " & cmb_GroupCharge.text & "_" & cmb_StoreCharge & "_" & txt_PaidPoint & "_" & txt_PaidPointMoney & ":�������� ����. �ٽ� �õ����ּ���(E00005)" & " " & Err.Description)
End Sub

'�˻�
Private Sub SSCommand3_Click()
    'MousePointer = vbHourglass '�𷡽ð���
    'Me.MousePointer = vbDefault '�⺻��
    sButton_Search = "����Ʈ��ȸ"
    Call ListView_PointDraw
End Sub

'��������
Private Sub ListView_PointDraw()
    Dim Column_to_size As Integer
    Dim rs As Recordset
    Dim rs2 As Recordset
    Dim sQry As String
    Dim sQry2 As String
    Dim itmX As ListItem
    Dim INDEX_NO As Long
    Dim bQryResult As Boolean
    Dim sStartDT, sEndDT As String
    Dim sIDSEQ As String
    Dim sID As String
    Dim sIDGUBUN As String
    Dim sPcode As String
    Dim i As Integer
    Dim strLine() As String
    Dim sStoreName As String
    
    Dim nRecordCount As String
    Dim nSumFreePoint As Long
    Dim nSumPaidPoint As Long
    Dim nSumPaidPointMoney As Long
    Dim nSumAutoFree As Long

'On Error GoTo Err_p

    Call ListViewExtended(ListView1)
    ListView1.View = lvwReport
    ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , " No "
    ListView1.ColumnHeaders.Add , , " ����      "
    ListView1.ColumnHeaders.Add , , " ���̵�          "
    ListView1.ColumnHeaders.Add , , " ��ü��                  "
    ListView1.ColumnHeaders.Add , , " ��������Ʈ  "
    ListView1.ColumnHeaders.Add , , " ��������Ʈ  "
    ListView1.ColumnHeaders.Add , , " �ݾ�                "
    'ListView1.ColumnHeaders.Add , , " ��������Ʈ �ڵ�����(�ſ�1��) "
    ListView1.ColumnHeaders.Add , , " ó������  "

    For Column_to_size = 0 To ListView1.ColumnHeaders.Count - 1
         SendMessage ListView1.hwnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
    Next


    sStartDT = Format(DTPicker1, "yyyy-mm-dd") & " 00:00:00"
    sEndDT = Format(DTPicker2, "yyyy-mm-dd") & " 23:59:59"
 
    If (cmb_GroupCharge.text = "��ü") Then
        sQry = "SELECT * FROM tb_id WHERE GUBUN != '�Ѱ�������' AND GUBUN != '������' AND GUBUN != '���'" '��� �׷�(���常 �ش�)
    Else
        sQry = "SELECT * FROM tb_id WHERE GUBUN = '" & cmb_GroupCharge.text & "' " 'Ư�� �׷�
    End If
    
    nSumFreePoint = 0
    nSumPaidPoint = 0
    nSumPaidPointMoney = 0
    nSumAutoFree = 0
    INDEX_NO = 1
    
    Set rs = New ADODB.Recordset
    rs.Open sQry, adoConn
    Do While Not (rs.EOF)
        
        sIDSEQ = rs!SEQ
        sID = rs!ID
        sIDGUBUN = rs!Gubun
        
        sQry2 = "SELECT * from tb_partner_log WHERE PCODE = '" & sIDSEQ & "' AND REG_DATE >= '" & sStartDT & "' AND REG_DATE <= '" & sEndDT & "' ORDER BY REG_DATE"
        Set rs2 = New ADODB.Recordset
        rs2.Open sQry2, adoConn

        Do While Not (rs2.EOF)
        
                If (cmb_StoreCharge.text = "��ü") Then

                    Set itmX = ListView1.ListItems.Add(, , "" & INDEX_NO)
                    
                    i = 1
                    itmX.SubItems(i) = "" & sIDGUBUN: i = i + 1
                    itmX.SubItems(i) = "" & sID: i = i + 1
                    sStoreName = GetStoreName(sIDSEQ)
                    itmX.SubItems(i) = "" & sStoreName: i = i + 1
                    itmX.SubItems(i) = "" & rs2!FREE_POINT: i = i + 1
                    itmX.SubItems(i) = "" & rs2!PAID_POINT: i = i + 1
                    itmX.SubItems(i) = "" & rs2!PAID_POINT_CHARGEMONEY: i = i + 1
                    'itmX.SubItems(i) = "" & rs!AUTOFREECHARGE: i = i + 1 '��������Ʈ �ڵ����� ������
                    itmX.SubItems(i) = "" & rs2!REG_DATE: i = i + 1

                    INDEX_NO = INDEX_NO + 1
                    If (Not rs2!PAID_POINT) Then nSumPaidPoint = nSumPaidPoint + rs2!PAID_POINT
                    If (Not rs2!FREE_POINT) Then nSumFreePoint = nSumFreePoint + rs2!FREE_POINT
                    If (Not rs2!PAID_POINT_CHARGEMONEY) Then nSumPaidPointMoney = nSumPaidPointMoney + rs2!PAID_POINT_CHARGEMONEY
                    'If (Not rs!AUTOFREECHARGE) Then nSumAutoFree = nSumAutoFree + rs!AUTOFREECHARGE
                    
                Else
                    strLine = Split(cmb_StoreCharge.text, ".")
                    sPcode = strLine(0)
                    
                    If (sIDSEQ = sPcode) Then
                        Set itmX = ListView1.ListItems.Add(, , "" & INDEX_NO)
                    
                        i = 1
                        itmX.SubItems(i) = "" & sIDGUBUN: i = i + 1
                        itmX.SubItems(i) = "" & sID: i = i + 1
                        itmX.SubItems(i) = "" & "��ü��": i = i + 1
                        itmX.SubItems(i) = "" & rs2!FREE_POINT: i = i + 1
                        itmX.SubItems(i) = "" & rs2!PAID_POINT: i = i + 1
                        itmX.SubItems(i) = "" & rs2!PAID_POINT_CHARGEMONEY: i = i + 1
                        'itmX.SubItems(i) = "" & rs!AUTOFREECHARGE: i = i + 1 '��������Ʈ �ڵ����� ������
                        itmX.SubItems(i) = "" & rs2!REG_DATE: i = i + 1
        
                        INDEX_NO = INDEX_NO + 1
                        If (Not rs2!PAID_POINT) Then nSumPaidPoint = nSumPaidPoint + rs2!PAID_POINT
                        If (Not rs2!FREE_POINT) Then nSumFreePoint = nSumFreePoint + rs2!FREE_POINT
                        If (Not rs2!PAID_POINT_CHARGEMONEY) Then nSumPaidPointMoney = nSumPaidPointMoney + rs2!PAID_POINT_CHARGEMONEY
                        'If (Not rs!AUTOFREECHARGE) Then nSumAutoFree = nSumAutoFree + rs!AUTOFREECHARGE
                    Else
                        'pass
                    End If
                    
                End If
                
                rs2.MoveNext
        Loop
        Set rs2 = Nothing
        
        rs.MoveNext
    Loop
    Set rs = Nothing
    
    Call PrintResult(INDEX_NO - 1, nSumFreePoint, nSumPaidPoint, nSumPaidPointMoney, nSumAutoFree) '���
    
Exit Sub
End Sub

Private Sub PrintResult(nCount As Long, nSumFree As Long, nSumPaid As Long, nSumMoney As Long, nAutoFree As Long)
    lbl_COUNT = "��ȸ�Ǽ�:" & nCount
    lbl_option(0) = "��������Ʈ : " & nSumFree '��������Ʈ
    lbl_option(1) = "��������Ʈ : " & nSumPaid '��������Ʈ
    lbl_option(2) = "�ݾ�:" & Format(nSumMoney, "###,###,##0") & " (��)" '�ݾ�
    lbl_option(3) = "��������Ʈ �ڵ����� : :" & nAutoFree  '��������Ʈ �ڵ�����
End Sub

Private Function GetStoreName(sSEQ As String)
    Dim rs As Recordset
    Dim sQry As String
    
    sQry = "SELECT PNAME from tb_partner WHERE SEQ = '" & sSEQ & "'"
    Set rs = New ADODB.Recordset
    rs.Open sQry, adoConn

    If Not (rs.EOF) Then
        GetStoreName = rs!PNAME
    Else
        GetStoreName = ""
    End If
    
    Set rs = Nothing
    
End Function
Private Sub Form_Load()
    'Dim Record_Source As String
    'Dim i As Integer
    
'On Error GoTo err_P
    
    Left = (Screen.width - width) / 2   ' ���� ���η� �߾ӿ� �����ϴ�.
    Top = (Screen.height - height) / 2   ' ���� ���η� �߾ӿ� �����ϴ�.
    
    
    DTPicker1.value = Now
    DTPicker2.value = Now
    
    Call Init_Charge
    
    sButton_Search = "��������ȸ"
    
Exit Sub
    
Err_P:
    MsgBox "������ ���̽� �������" & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & "��Ʈ�� �����ڿ��� ���� �ٶ��ϴ�." & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & "������ ���̽� ���������� �ڷ�˻� ����� �����Ҽ� �����ϴ�.", vbCritical
End Sub

Private Sub Init_Charge()
    
    Call SetCombo_ChargeGroup
    Call SetCombo_ChargeStore("��ü")
    
    txt_PaidPoint.text = "0"
    txt_PaidPointMoney = "0"
    txt_FreePoint.text = "0"
    
End Sub


Private Sub SetCombo_ChargeGroup()
    Dim rs As Recordset
    Dim sQry As String
    
On Error GoTo Err_P

    cmb_GroupCharge.Clear
    cmb_GroupCharge.AddItem "��ü"
    
    sQry = "SELECT GUBUN From tb_id Group By GUBUN"
    Set rs = New ADODB.Recordset
    rs.Open sQry, adoConn
    
    If Not rs.EOF Then
        Do While Not (rs.EOF)
            If (rs!Gubun <> "�Ѱ�������" And rs!Gubun <> "������" And rs!Gubun <> "���") Then
                cmb_GroupCharge.AddItem "" & rs!Gubun
            End If
            rs.MoveNext
        Loop
    End If
    Set rs = Nothing

    cmb_GroupCharge.ListIndex = 0
    
Exit Sub
Err_P:
    Call DataLogger("[FrmWebdc SetCombo_ChargeGroup]    " & Err.Description & " " & sQry)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Unload Me
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

    If (sButton_Search = "����Ʈ��ȸ") Then

        Dim nFree As Long
        Dim nPaid As Long
        Dim nMoney As Long
        Dim nAuto As Long
        
        With ListView1.SelectedItem
            If (("" & .SubItems(4)) = "") Then
                nFree = 0
            Else
                nFree = .SubItems(4)
            End If
            If (("" & .SubItems(5)) = "") Then
                nPaid = 0
            Else
                nPaid = .SubItems(5)
            End If
            
            If (("" & .SubItems(6)) = "") Then
                nMoney = 0
            Else
                nMoney = .SubItems(6)
            End If
            
            If (("" & .SubItems(7)) = "") Then
                nAuto = 0
            Else
                nAuto = .SubItems(7)
            End If
            
            Call PrintResult(1, nFree, nPaid, nMoney, nAuto) '���
        End With
        
    Else
        Call PrintResult(0, 0, 0, 0, 0) '���
    End If
    
End Sub

Private Sub SSCommand1_Click()

    Dim tmpFileName As String

On Error GoTo Err_P
    tmpFileName = Format(Now, "YYYYMMDD_HHMMSS")
    tmpFileName = App.Path & "\Excel\" & tmpFileName & "_" & sButton_Search
        
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

'����
Private Sub SSCommand2_Click()
    Unload Me
End Sub

'������ ��ȸ
Private Sub SSCommand4_Click()
    sButton_Search = "��������ȸ"
    Call ListView_WebdcDraw
    Call PrintResult(0, 0, 0, 0, 0) '���
End Sub

Private Sub ListView_WebdcDraw()
    Dim Column_to_size As Integer
    Dim rs As Recordset
    Dim rs2 As Recordset
    Dim sQry As String
    Dim sQry2 As String
    Dim itmX As ListItem
    Dim INDEX_NO As Long
    Dim bQryResult As Boolean
    Dim sStartDT, sEndDT As String
    Dim sIDSEQ As String
    Dim sID As String
    Dim sIDGUBUN As String
    Dim sPcode As String
    Dim i As Integer
    Dim strLine() As String
    Dim sStoreName As String
'    Dim sPcode As String
    Dim nRecordCount As String
    Dim nSumFreePoint As Long
    Dim nSumPaidPoint As Long
    Dim nSumPaidPointMoney As Long
    Dim nSumAutoFree As Long
    

'On Error GoTo Err_p

    Call ListViewExtended(ListView1)
    ListView1.View = lvwReport
    ListView1.ListItems.Clear
    ListView1.ColumnHeaders.Clear
    ListView1.ColumnHeaders.Add , , " No "
    ListView1.ColumnHeaders.Add , , " ��ü��      "
    ListView1.ColumnHeaders.Add , , " ������ȣ      "
    ListView1.ColumnHeaders.Add , , " ����          "
    ListView1.ColumnHeaders.Add , , " ����Ʈ                  "
    ListView1.ColumnHeaders.Add , , " ó���Ͻ�  "

    For Column_to_size = 0 To ListView1.ColumnHeaders.Count - 1
         SendMessage ListView1.hwnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
    Next

    cmb_GroupCharge.ListIndex = 0
    
    sStartDT = Format(DTPicker1, "yyyy-mm-dd") & " 00:00:00"
    sEndDT = Format(DTPicker2, "yyyy-mm-dd") & " 23:59:59"
    
    If (cmb_StoreCharge.text = "��ü") Then
        sQry = "SELECT * FROM tb_web_dc WHERE DT_DC >= '" & sStartDT & "' AND DT_DC <= '" & sEndDT & "'  "
    Else
        strLine = Split(cmb_StoreCharge.text, ".")
        sPcode = strLine(0)
        sQry = "SELECT * FROM tb_web_dc WHERE PCODE = '" & sPcode & "' AND DT_DC >= '" & sStartDT & "' AND DT_DC <= '" & sEndDT & "' "
    End If
    
    nSumFreePoint = 0
    nSumPaidPoint = 0
    nSumPaidPointMoney = 0
    nSumAutoFree = 0
    INDEX_NO = 1
    
    Set rs = New ADODB.Recordset
    rs.Open sQry, adoConn
    Do While Not (rs.EOF)
        
        Set itmX = ListView1.ListItems.Add(, , "" & INDEX_NO)

        i = 1
        itmX.SubItems(i) = "" & rs!PNAME: i = i + 1   '��ü��
        itmX.SubItems(i) = "" & rs!DC_CARNO: i = i + 1 '������ȣ
        itmX.SubItems(i) = "" & rs!DC_NAME: i = i + 1  '���γ���
        itmX.SubItems(i) = "-" & rs!DC_CODE: i = i + 1  '�������Ʈ
        itmX.SubItems(i) = "" & Format(rs!DT_DC, "yyyy-mm-dd hh:nn:ss"): i = i + 1    'ó���Ͻ�

        INDEX_NO = INDEX_NO + 1
'        If (Not rs2!PAID_POINT) Then nSumPaidPoint = nSumPaidPoint + rs2!PAID_POINT
'        If (Not rs2!FREE_POINT) Then nSumFreePoint = nSumFreePoint + rs2!FREE_POINT
'        If (Not rs2!PAID_POINT_CHARGEMONEY) Then nSumPaidPointMoney = nSumPaidPointMoney + rs2!PAID_POINT_CHARGEMONEY
'        If (Not rs!AUTOFREECHARGE) Then nSumAutoFree = nSumAutoFree + rs!AUTOFREECHARGE
                
        rs.MoveNext
    Loop

    Set rs = Nothing
    
    'Call PrintResult(INDEX_NO - 1, nSumFreePoint, nSumPaidPoint, nSumPaidPointMoney, nSumAutoFree) '���
    
Exit Sub
End Sub



Private Sub txt_FreePoint_KeyPress(KeyAscii As Integer)
    '�������Է�
    If (txt_FreePoint = "0") Then
        txt_FreePoint = ""
    End If

    If (KeyAscii = 45) Then
        txt_FreePoint = ""
    ElseIf (KeyAscii = vbKeyBack Or (KeyAscii >= vbKey0 And KeyAscii <= vbKey9)) Then '�齺���̽�, ����
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txt_PaidPoint_KeyPress(KeyAscii As Integer)
    '�������Է�
    If (txt_PaidPoint = "0") Then
        txt_PaidPoint = ""
    End If

    If (KeyAscii = 45) Then
        txt_PaidPoint = ""
    ElseIf (KeyAscii = vbKeyBack Or (KeyAscii >= vbKey0 And KeyAscii <= vbKey9)) Then '�齺���̽�, ����
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub txt_PaidPointMoney_KeyPress(KeyAscii As Integer)
    '�������Է�
    If (txt_PaidPointMoney = "0") Then
        txt_PaidPointMoney = ""
    End If

    If (KeyAscii = 45) Then
        txt_PaidPointMoney = ""
    ElseIf (KeyAscii = vbKeyBack Or (KeyAscii >= vbKey0 And KeyAscii <= vbKey9)) Then '�齺���̽�, ����
    Else
        KeyAscii = 0
    End If
End Sub


