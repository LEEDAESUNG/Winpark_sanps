VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form FrmGateBar 
   Caption         =   "FrmGateBar"
   ClientHeight    =   3810
   ClientLeft      =   10530
   ClientTop       =   3720
   ClientWidth     =   6780
   BeginProperty Font 
      Name            =   "나눔고딕"
      Size            =   9.75
      Charset         =   129
      Weight          =   600
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3810
   ScaleWidth      =   6780
   Begin VB.CommandButton cmd_Exit 
      Caption         =   "Cancel"
      Height          =   570
      Left            =   5550
      TabIndex        =   1
      Top             =   150
      Width           =   1095
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   1245
      Index           =   0
      Left            =   2505
      TabIndex        =   0
      Top             =   1935
      Width           =   1725
      _Version        =   65536
      _ExtentX        =   3043
      _ExtentY        =   2196
      _StockProps     =   78
      Caption         =   "OPEN"
      ForeColor       =   32768
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   20.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   1245
      Index           =   1
      Left            =   7530
      TabIndex        =   3
      Top             =   1935
      Width           =   1725
      _Version        =   65536
      _ExtentX        =   3043
      _ExtentY        =   2196
      _StockProps     =   78
      Caption         =   "CLOSE"
      ForeColor       =   192
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "나눔고딕"
         Size            =   20.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSWinsockLib.Winsock Gate_Winsock 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label lbl_GateName 
      Alignment       =   2  '가운데 맞춤
      Caption         =   "GateName"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   26.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   630
      Left            =   375
      TabIndex        =   4
      Top             =   840
      Width           =   5970
   End
   Begin VB.Label lbl_lpIP 
      Caption         =   "192.168.123.123"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   120
      Width           =   2220
   End
End
Attribute VB_Name = "FrmGateBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tmpDis1, tmpDis2 As String * 32
Dim LprQry, CMD, CMD_IP As String



Private Sub Form_Load()
    Dim i As Integer
    Dim rs As Recordset
    Dim Qry As String

    Left = (Screen.Width - Width) / 2   ' 폼을 가로로 중앙에 놓습니다.
    Top = (Screen.Height - Height) / 2   ' 폼을 세로로 중앙에 놓습니다.
    
    Qry = "SELECT * From TB_LPR Where IP = '" & Glo_GateBar_IP & "'"
    Set rs = New ADODB.Recordset
    rs.Open Qry, adoConn
    
    If (rs.EOF) Then
        Set rs = Nothing
        lbl_GateName.Caption = "해당 LPR이 없습니다..!!"
    Else
        lbl_lpIP.Caption = Glo_GateBar_IP
        lbl_GateName.Caption = "" & rs!GateName
    End If
    Set rs = Nothing
    
    Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    차량등록/관리 시작...!!")

End Sub

Private Sub cmd_Button_Click(Index As Integer)
    Dim i, j As Integer
    Dim myExcelFile As New ExcelFile
    Dim tmpFileName As String

    Select Case Index
        Case 0
            If (Len(Glo_GateBar_IP) <> 0) Then
                CMD = ""
                CMD = "CMD_RELAY_01"
                Call Socket_ConnectGate(Glo_GateBar_IP, 233)
            End If
            Exit Sub
        Case 1
            If (Len(Glo_GateBar_IP) <> 0) Then
                CMD = ""
                CMD = "CMD_RELAY_02"
                Call Socket_ConnectGate(Glo_GateBar_IP, 233)
            End If
            Exit Sub
    End Select
End Sub

'취소 버튼
Private Sub cmd_Exit_Click()
    Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    GateBar CMD Cancel..!!")
    Unload Me
End Sub


Private Sub Socket_ConnectGate(ByVal IP As String, ByVal Port As Long)
    'Gate_Winsock.Close

    If (Gate_Winsock.State <> sckClosed) Then
        Gate_Winsock.Close
        'DoEvents
    End If
    Gate_Winsock.Connect IP, Port

    'Call sOutput("[Gate 접속]", CMD & "  " & IP)
    'Call Err_doc("    [Gate 접속]  시도 IP = " & IP & "    PORT = " & Port)
End Sub

Private Sub Gate_Winsock_Connect()
    Dim bData() As Byte

    ReDim bData(Len(CMD) - 1) As Byte
    bData = StrConv(CMD, vbFromUnicode)
    Gate_Winsock.SendData bData

    'Call sOutput("[Gate 송신]", CMD)
'    If (Check5.value = 1) Then
'        Call Err_doc("    [Gate 송신] " & CMD)
'    End If
    'Fee_sock.Close
End Sub

Private Sub Gate_Winsock_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String

    Gate_Winsock.GetData strData, , bytesTotal
    'Call sOutput("[Gate 수신]", strData)
    Gate_Winsock.Close
    
    Unload Me
End Sub

Private Sub Gate_Winsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'Call sOutput(Source, "[Gate 소켓] " & "에러 : " & Description)
'    If (Check5.value = 1) Then
'        Call Err_doc("   [Gate 소켓] " & "에러 : " & Description)
'    End If
End Sub
