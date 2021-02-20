VERSION 5.00
Begin VB.Form FrmUnita 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  '단일 고정
   Caption         =   "ParkingManager™"
   ClientHeight    =   10125
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   20220
   LinkTopic       =   "FrmUnita"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10125
   ScaleMode       =   0  '사용자
   ScaleWidth      =   19206.01
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   1455
      Left            =   -30
      TabIndex        =   22
      Top             =   -90
      Width           =   20280
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "통신방법"
         Height          =   885
         Left            =   300
         TabIndex        =   25
         Top             =   360
         Width           =   2415
         Begin VB.ComboBox cmb_DeviceMode 
            Height          =   300
            Index           =   0
            ItemData        =   "FrmUnita.frx":0000
            Left            =   180
            List            =   "FrmUnita.frx":000A
            Style           =   2  '드롭다운 목록
            TabIndex        =   26
            Top             =   420
            Width           =   2070
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "저장"
         Height          =   405
         Left            =   15540
         TabIndex        =   24
         Top             =   525
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "닫기"
         Height          =   405
         Left            =   17130
         TabIndex        =   23
         Top             =   525
         Width           =   1455
      End
   End
   Begin VB.CheckBox chk_UseYN 
      BackColor       =   &H00E0E0E0&
      Caption         =   "사용"
      Height          =   210
      Index           =   0
      Left            =   2760
      TabIndex        =   20
      Top             =   2055
      Width           =   660
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "LANE1"
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   5115
      Index           =   0
      Left            =   45
      TabIndex        =   1
      Top             =   2190
      Width           =   3375
      Begin VB.ComboBox cmb_Protocol 
         Height          =   300
         Index           =   0
         ItemData        =   "FrmUnita.frx":0018
         Left            =   1230
         List            =   "FrmUnita.frx":0022
         Style           =   2  '드롭다운 목록
         TabIndex        =   21
         Top             =   1815
         Width           =   1725
      End
      Begin VB.TextBox txt_IP 
         Height          =   330
         Index           =   0
         Left            =   1230
         TabIndex        =   14
         Text            =   "192.168.0.222"
         Top             =   675
         Width           =   1515
      End
      Begin VB.ComboBox cmb_Model 
         Height          =   300
         Index           =   0
         ItemData        =   "FrmUnita.frx":0030
         Left            =   1230
         List            =   "FrmUnita.frx":0032
         Style           =   2  '드롭다운 목록
         TabIndex        =   13
         Top             =   360
         Width           =   1725
      End
      Begin VB.TextBox txt_Gateway 
         Height          =   330
         Index           =   0
         Left            =   1230
         TabIndex        =   12
         Text            =   "192.168.0.1"
         Top             =   1035
         Width           =   1515
      End
      Begin VB.TextBox txt_Subnetmask 
         Height          =   330
         Index           =   0
         Left            =   1230
         TabIndex        =   11
         Text            =   "24"
         Top             =   1395
         Width           =   555
      End
      Begin VB.CommandButton cmd_GateLock 
         Caption         =   "차단기잠금"
         Enabled         =   0   'False
         Height          =   435
         Index           =   0
         Left            =   405
         TabIndex        =   10
         ToolTipText     =   "차단기 오픈명령 전송"
         Top             =   3300
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton cmd_GateReset 
         Caption         =   "차단기리셋"
         Height          =   435
         Index           =   0
         Left            =   1335
         TabIndex        =   9
         ToolTipText     =   "차단기 오픈명령 전송"
         Top             =   3300
         Width           =   735
      End
      Begin VB.CommandButton cmd_GateUp 
         Caption         =   "차단기Up"
         Height          =   435
         Index           =   0
         Left            =   405
         TabIndex        =   8
         ToolTipText     =   "차단기 오픈명령 전송"
         Top             =   3870
         Width           =   735
      End
      Begin VB.CommandButton cmd_SoundSpecial 
         Caption         =   "사운드 (특수)"
         Height          =   435
         Index           =   0
         Left            =   2265
         TabIndex        =   7
         Top             =   4440
         Width           =   735
      End
      Begin VB.CommandButton cmd_SoundWarning 
         Caption         =   "사운드 (경고)"
         Height          =   435
         Index           =   0
         Left            =   1335
         TabIndex        =   6
         Top             =   4440
         Width           =   735
      End
      Begin VB.CommandButton cmd_SoundNormal 
         Caption         =   "사운드 (일반)"
         Height          =   435
         Index           =   0
         Left            =   405
         TabIndex        =   5
         Top             =   4440
         Width           =   735
      End
      Begin VB.CommandButton cmd_GateStatus 
         Caption         =   "차단기Status"
         Height          =   435
         Index           =   0
         Left            =   2265
         TabIndex        =   4
         ToolTipText     =   "차단기 상태 요청"
         Top             =   3870
         Width           =   735
      End
      Begin VB.CommandButton cmd_GateDown 
         Caption         =   "차단기Down"
         Height          =   435
         Index           =   0
         Left            =   1335
         TabIndex        =   3
         ToolTipText     =   "차단기 다운명령 전송"
         Top             =   3870
         Width           =   735
      End
      Begin VB.CommandButton cmd_Reboot 
         Caption         =   "리부팅"
         Height          =   435
         Index           =   0
         Left            =   2265
         TabIndex        =   2
         ToolTipText     =   "유니타 장치 리부팅합니다"
         Top             =   3300
         Width           =   735
      End
      Begin VB.Label lbl_IP 
         BackColor       =   &H00E0E0E0&
         Caption         =   "IP"
         Height          =   165
         Index           =   0
         Left            =   315
         TabIndex        =   19
         Top             =   780
         Width           =   585
      End
      Begin VB.Label lbl_Model 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Model"
         Height          =   165
         Index           =   0
         Left            =   315
         TabIndex        =   18
         Top             =   420
         Width           =   585
      End
      Begin VB.Label lbl_Gateway 
         BackColor       =   &H00E0E0E0&
         Caption         =   "G/W"
         Height          =   165
         Index           =   0
         Left            =   315
         TabIndex        =   17
         Top             =   1140
         Width           =   585
      End
      Begin VB.Label lbl_Subnetmask 
         BackColor       =   &H00E0E0E0&
         Caption         =   "S/M"
         Height          =   165
         Index           =   0
         Left            =   315
         TabIndex        =   16
         Top             =   1500
         Width           =   585
      End
      Begin VB.Label lbl_Protocol 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Protocol"
         Height          =   165
         Index           =   0
         Left            =   315
         TabIndex        =   15
         Top             =   1860
         Width           =   675
      End
   End
   Begin VB.ListBox ListData 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   2580
      Left            =   0
      TabIndex        =   0
      Top             =   7590
      Width           =   20235
   End
End
Attribute VB_Name = "FrmUnita"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chk_Unita_YN_Click(Index As Integer)
    If (chk_Unita_YN(Index).value = 1) Then
        Frame2(Index).Enabled = True
        txt_UnitaIP(Index).Enabled = True
        txt_UnitaGW(Index).Enabled = True
        txt_UnitaSM(Index).Enabled = True
        txt_UnitaDNS(Index).Enabled = True
        txt_UnitaLoop(Index).Enabled = True
        txt_UnitaGate(Index).Enabled = True
        cmb_UnitaModel(Index).Enabled = True
        txt_UnitaActiveTime(Index).Enabled = True
        txt_UnitaRelay_Time(Index).Enabled = True
        txt_UnitaCPUtemperature(Index).Enabled = True
'        cmd_Unita_Reboot(Index).Enabled = True
'        cmd_Unita_ReSet(Index).Enabled = True
    Else
        Frame2(Index).Enabled = False
        txt_UnitaIP(Index).Enabled = False
        txt_UnitaGW(Index).Enabled = False
        txt_UnitaSM(Index).Enabled = False
        txt_UnitaDNS(Index).Enabled = False
        txt_UnitaLoop(Index).Enabled = False
        txt_UnitaGate(Index).Enabled = False
        cmb_UnitaModel(Index).Enabled = False
        txt_UnitaActiveTime(Index).Enabled = False
        txt_UnitaRelay_Time(Index).Enabled = False
        txt_UnitaCPUtemperature(Index).Enabled = False
'        cmd_Unita_Reboot(Index).Enabled = False
'        cmd_Unita_ReSet(Index).Enabled = False
        
    End If
End Sub

Private Sub cmb_UnitaModel_Change(Index As Integer)

        If cmb_UnitaModel(Index).name = "ELEPARTS-3RELAY-BOARD" Then
            Call SetRPiControl(Index, False)
            Call SetElepartsControl(Index, True)
            
        ElseIf cmb_UnitaModel(Index) = "RPi-3RELAY-BOARD" Then
            Call SetElepartsControl(Index, False)
            Call SetRPiControl(Index, True)
        End If
    
End Sub

Private Sub SetRPiControl(Index As Integer, val As Boolean)
    lbl_UnitaIP(Index).Visible = val
    txt_UnitaIP(Index).Enabled = val
    txt_UnitaIP(Index).Visible = val
    
    lbl_UnitaGW(Index).Visible = val
    lbl_UnitaGW(Index).Enabled = val
    txt_UnitaGW(Index).Visible = val
    txt_UnitaGW(Index).Enabled = val
    
    lbl_UnitaSM(Index).Visible = val
    lbl_UnitaSM(Index).Enabled = val
    txt_UnitaSM(Index).Visible = val
    txt_UnitaSM(Index).Enabled = val
    
    
    lbl_UnitaDNS(Index).Visible = val
    lbl_UnitaDNS(Index).Enabled = val
    txt_UnitaDNS(Index).Visible = val
    txt_UnitaDNS(Index).Enabled = val
    
    lbl_UnitaLoop(Index).Visible = False
    lbl_UnitaLoop(Index).Enabled = False
    txt_UnitaLoop(Index).Visible = False
    txt_UnitaLoop(Index).Enabled = False
    
    lbl_UnitaGate(Index).Visible = False
    lbl_UnitaGate(Index).Enabled = False
    txt_UnitaGate(Index).Visible = False
    txt_UnitaGate(Index).Enabled = False
    
    lbl_UnitaActiveTime(Index).Visible = val
    lbl_UnitaActiveTime(Index).Enabled = val
    txt_UnitaActiveTime(Index).Visible = val
    txt_UnitaActiveTime(Index).Enabled = val
    lbl_UnitaActiveTime_s(Index).Visible = val
    lbl_UnitaActiveTime_s(Index).Enabled = val
    
    lbl_UnitaRelay_Time(Index).Visible = False
    lbl_UnitaRelay_Time(Index).Enabled = False
    txt_UnitaRelay_Time(Index).Visible = False
    txt_UnitaRelay_Time(Index).Enabled = False
    lbl_UnitaRelay_Time_s(Index).Visible = False
    lbl_UnitaRelay_Time_s(Index).Enabled = False
    
    lbl_UnitaCPUtemperature(Index).Visible = val
    lbl_UnitaCPUtemperature(Index).Enabled = val
    txt_UnitaCPUtemperature(Index).Visible = val
    txt_UnitaCPUtemperature(Index).Enabled = val
    lbl_UnitaCPUtemperature_c(Index).Visible = val
    lbl_UnitaCPUtemperature_c(Index).Enabled = val

    cmd_Unitatemp(Index).Visible = val
    cmd_Unitatemp(Index).Enabled = val
    
End Sub


Private Sub SetElepartsControl(Index As Integer, val As Boolean)
    lbl_UnitaIP(Index).Visible = val
    txt_UnitaIP(Index).Enabled = val
    txt_UnitaIP(Index).Visible = val
                
    lbl_UnitaGW(Index).Visible = val
    lbl_UnitaGW(Index).Enabled = val
    txt_UnitaGW(Index).Visible = val
    txt_UnitaGW(Index).Enabled = val
    
    lbl_UnitaSM(Index).Visible = val
    lbl_UnitaSM(Index).Enabled = val
    txt_UnitaSM(Index).Visible = val
    txt_UnitaSM(Index).Enabled = val
    
    
    lbl_UnitaDNS(Index).Visible = val
    lbl_UnitaDNS(Index).Enabled = val
    txt_UnitaDNS(Index).Visible = val
    txt_UnitaDNS(Index).Enabled = val
    
    lbl_UnitaLoop(Index).Visible = val
    lbl_UnitaLoop(Index).Enabled = val
    txt_UnitaLoop(Index).Visible = val
    txt_UnitaLoop(Index).Enabled = val
    
    lbl_UnitaGate(Index).Visible = val
    lbl_UnitaGate(Index).Enabled = val
    txt_UnitaGate(Index).Visible = val
    txt_UnitaGate(Index).Enabled = val
    
    lbl_UnitaActiveTime(Index).Visible = val
    lbl_UnitaActiveTime(Index).Enabled = val
    txt_UnitaActiveTime(Index).Visible = val
    txt_UnitaActiveTime(Index).Enabled = val
    lbl_UnitaActiveTime_s(Index).Visible = val
    lbl_UnitaActiveTime_s(Index).Enabled = val
    
    lbl_UnitaRelay_Time(Index).Visible = val
    lbl_UnitaRelay_Time(Index).Enabled = val
    txt_UnitaRelay_Time(Index).Visible = val
    txt_UnitaRelay_Time(Index).Enabled = val
    lbl_UnitaRelay_Time_s(Index).Visible = val
    lbl_UnitaRelay_Time_s(Index).Enabled = val
    
    lbl_UnitaCPUtemperature(Index).Visible = val
    lbl_UnitaCPUtemperature(Index).Enabled = val
    txt_UnitaCPUtemperature(Index).Visible = val
    txt_UnitaCPUtemperature(Index).Enabled = val
    lbl_UnitaCPUtemperature_c(Index).Visible = val
    lbl_UnitaCPUtemperature_c(Index).Enabled = val

    cmd_Unitatemp(Index).Visible = val
    cmd_Unitatemp(Index).Enabled = val
    
End Sub

Private Sub cmd_Unita_GateDown_Click(Index As Integer)
    
    If (cmb_UnitaModel(Index).text = "ELEPARTS-3RELAY-BOARD") Then
        
    ElseIf (cmb_UnitaModel(Index).text = "RPi-3RELAY-BOARD") Then
        Call DataLogger("[Unita GateDown]  Target Gate = " & Index)
        Call Unita_Command_Send(Index, "GATE_DOWN")
    End If

End Sub

Private Sub cmd_Unita_GateLock_Click(Index As Integer)
'    If (cmb_UnitaModel(Index).Text = "ELEPARTS-3RELAY-BOARD") Then
'
'    ElseIf (cmb_UnitaModel(Index).Text = "RPi-3RELAY-BOARD") Then
'        Call DataLogger("[Unita GateReset]  Target Gate = " & Index)
'
'        If (cmd_Unita_GateLock(Index).Caption = "차단기잠금") Then
'            Call Unita_Command_Send(Index, "GATE_LOCK")
'            cmd_Unita_GateLock(Index).Caption = "차단기풀림"
'        Else
'            Call Unita_Command_Send(Index, "GATE_UNLOCK")
'            cmd_Unita_GateLock(Index).Caption = "차단기잠금"
'        End If
'
'    End If

End Sub

Private Sub cmd_Unita_GateReset_Click(Index As Integer)
    If (cmb_UnitaModel(Index).text = "ELEPARTS-3RELAY-BOARD") Then
        
    ElseIf (cmb_UnitaModel(Index).text = "RPi-3RELAY-BOARD") Then
        Call DataLogger("[Unita GateReset]  Target Gate = " & Index)
        Call Unita_Command_Send(Index, "GATE_RESET")
    End If

End Sub

Private Sub cmd_Unita_GateStatus_Click(Index As Integer)
    
    If (cmb_UnitaModel(Index).text = "ELEPARTS-3RELAY-BOARD") Then
        
    ElseIf (cmb_UnitaModel(Index).text = "RPi-3RELAY-BOARD") Then
        Call DataLogger("[Unita GateStatus]  Target Gate = " & Index)
        Call Unita_Command_Send(Index, "GATE_STATUS")
    End If
    
End Sub

Private Sub cmd_Unita_GateUp_Click(Index As Integer)

    If (cmb_UnitaModel(Index).text = "ELEPARTS-3RELAY-BOARD") Then
        
    ElseIf (cmb_UnitaModel(Index).text = "RPi-3RELAY-BOARD") Then
        Call DataLogger("[Unita GateUp]  Target Gate = " & Index)
        Call Unita_Command_Send(Index, "GATE_UP")
    End If
    
End Sub

Private Sub cmd_Unita_Reboot_Click(Index As Integer)

    If (cmb_UnitaModel(Index).text = "ELEPARTS-3RELAY-BOARD") Then
        Call DataLogger("[Unita Reboot]  Target Gate = " & Index)
        Call Unita_Command_Send(Index, "H_Reboot")
        
    ElseIf (cmb_UnitaModel(Index).text = "RPi-3RELAY-BOARD") Then
        Call DataLogger("[Unita Reboot]  Target Gate = " & Index)
        Call Unita_Command_Send(Index, "REBOOT")
    End If
    
End Sub

Private Sub cmd_Unita_ReSet_Click(Index As Integer)
    Dim sYN As String
    Dim sMODEL, sIP, sGW, sSM, sDNS, sLOOP, sGATE, sACTTIME, sRELAYTIME As String
    Dim preIP As String
    Dim prePORT As Long
    Dim cmd As String

    If (chk_Unita_YN(Index).value = 1) Then
        sYN = "Y"
    Else
        sYN = "N"
    End If

    If (cmb_UnitaModel(Index).text = "ELEPARTS-3RELAY-BOARD") Then
        sMODEL = UCase(cmb_UnitaModel(Index).text)
        sIP = Trim(txt_UnitaIP(Index).text)
        sGW = Trim(txt_UnitaGW(Index).text)
        sSM = Trim(txt_UnitaSM(Index).text)
        sDNS = Trim(txt_UnitaDNS(Index).text)
        sLOOP = Trim(txt_UnitaLoop(Index).text)
        sGATE = Trim(txt_UnitaGate(Index).text)
        sACTTIME = Trim(txt_UnitaActiveTime(Index).text)
        sRELAYTIME = Trim(txt_UnitaRelay_Time(Index).text)

        Call DataLogger("[Unita Reset]  Target Gate = " & Index)
        Call Unita_Command_Send(Index, "H_Reset_" & sYN & "_" & sMODEL & "_" & sIP & "_" & sGW & "_" & sSM & "_" & sDNS & "_" & sLOOP & "_" & sGATE & "_" & sACTTIME & "_" & sRELAYTIME)


    ElseIf (cmb_UnitaModel(Index).text = "RPi-3RELAY-BOARD") Then
        sMODEL = UCase(cmb_UnitaModel(Index).text)
        sIP = Trim(txt_UnitaIP(Index).text)
        sGW = Trim(txt_UnitaGW(Index).text)
        sSM = Trim(txt_UnitaSM(Index).text)
        sDNS = Trim(txt_UnitaDNS(Index).text)
        'sLOOP = Trim(txt_UnitaLoop(index).Text)
        'sGATE = Trim(txt_UnitaGate(index).Text)
        sACTTIME = Trim(txt_UnitaActiveTime(Index).text)
        'sRELAYTIME = Trim(txt_UnitaRelay_Time(index).Text)
        
        Call DataLogger("[Unita Use]  Target Gate = " & Index)
        Call Unita_Command_Send(Index, "USE_" & sYN)
        
        Call DataLogger("[Unita Reset]  Target Gate = " & Index)
        Call Unita_Command_Send(Index, "NWRESET_IP_" & sIP & "_GW_" & sGW & "_SM_" & sSM & "_DNS_" & sDNS)

    End If

End Sub


Private Sub cmd_Unita_SoundNormal_Click(Index As Integer)
    
    If (cmb_UnitaModel(Index).text = "ELEPARTS-3RELAY-BOARD") Then
        
    ElseIf (cmb_UnitaModel(Index).text = "RPi-3RELAY-BOARD") Then
        Call DataLogger("[Unita Sound Normal]  Target Gate = " & Index)
        Call Unita_Command_Send(Index, "SOUND_NORMAL")
    End If
    
End Sub

Private Sub cmd_Unita_SoundSpecial_Click(Index As Integer)
    
    If (cmb_UnitaModel(Index).text = "ELEPARTS-3RELAY-BOARD") Then
        
    ElseIf (cmb_UnitaModel(Index).text = "RPi-3RELAY-BOARD") Then
        Call DataLogger("[Unita Sound Special]  Target Gate = " & Index)
        Call Unita_Command_Send(Index, "SOUND_SPECIAL")
    End If
    
End Sub

Private Sub cmd_Unita_SoundWarning_Click(Index As Integer)
    If (cmb_UnitaModel(Index).text = "ELEPARTS-3RELAY-BOARD") Then
        
    ElseIf (cmb_UnitaModel(Index).text = "RPi-3RELAY-BOARD") Then
        Call DataLogger("[Unita Sound Warning]  Target Gate = " & Index)
        Call Unita_Command_Send(Index, "SOUND_WARNING")
    End If
    
End Sub

Private Sub cmd_UnitaDispSpeed_Click(Index As Integer)
    Dim ip As String
    Dim Port As Long
    
    Call DataLogger("[Unita Display Speed]  Target Gate = " & Index)
    Call Unita_Command_Send(Index, "SPUP_1000_SPDN_1000")
End Sub

Private Sub cmd_UnitaDownDispSpeed_Click(Index As Integer)

End Sub

Private Sub cmd_Unitatemp_Click(Index As Integer)
    Dim ip As String
    Dim Port As Long
    
    
    If (cmb_UnitaModel(Index).text = "ELEPARTS-3RELAY-BOARD") Then
        Select Case Index
            Case 0
                Unita_Cmd_Str(Index) = "H_Info_CPUtemperature"
                FrmTcpServer.Unita1_cmd_sock(Index).Close
                FrmTcpServer.Unita1_cmd_sock(Index).Protocol = sckTCPProtocol
                FrmTcpServer.Unita1_cmd_sock(Index).Connect Trim(LANE1_UnitaIP), LANE1_UnitaPort
            Case 1
                Unita_Cmd_Str(Index) = "H_Info_CPUtemperature"
                FrmTcpServer.Unita1_cmd_sock(Index).Close
                FrmTcpServer.Unita1_cmd_sock(Index).Protocol = sckTCPProtocol
                FrmTcpServer.Unita1_cmd_sock(Index).Connect Trim(LANE2_UnitaIP), LANE2_UnitaPort
            Case 2
                Unita_Cmd_Str(Index) = "H_Info_CPUtemperature"
                FrmTcpServer.Unita1_cmd_sock(Index).Close
                FrmTcpServer.Unita1_cmd_sock(Index).Protocol = sckTCPProtocol
                FrmTcpServer.Unita1_cmd_sock(Index).Connect Trim(LANE3_UnitaIP), LANE3_UnitaPort
            Case 3
                Unita_Cmd_Str(Index) = "H_Info_CPUtemperature"
                FrmTcpServer.Unita1_cmd_sock(Index).Close
                FrmTcpServer.Unita1_cmd_sock(Index).Protocol = sckTCPProtocol
                FrmTcpServer.Unita1_cmd_sock(Index).Connect Trim(LANE4_UnitaIP), LANE4_UnitaPort
            Case 4
                Unita_Cmd_Str(Index) = "H_Info_CPUtemperature"
                FrmTcpServer.Unita1_cmd_sock(Index).Close
                FrmTcpServer.Unita1_cmd_sock(Index).Protocol = sckTCPProtocol
                FrmTcpServer.Unita1_cmd_sock(Index).Connect Trim(LANE5_UnitaIP), LANE5_UnitaPort
            Case 5
                Unita_Cmd_Str(Index) = "H_Info_CPUtemperature"
                FrmTcpServer.Unita1_cmd_sock(Index).Close
                FrmTcpServer.Unita1_cmd_sock(Index).Protocol = sckTCPProtocol
                FrmTcpServer.Unita1_cmd_sock(Index).Connect Trim(LANE6_UnitaIP), LANE6_UnitaPort
        End Select
        
    ElseIf (cmb_UnitaModel(Index).text = "RPi-3RELAY-BOARD") Then
        Call DataLogger("[Unita TEMPERATURE check]  Target Gate = " & Index)
        Call Unita_Command_Send(Index, "TEMPERATURE")
        
    Else
        Call DataLogger("[Unita TEMPERATURE check]  Error Model: " & cmb_UnitaModel(Index).text)
    End If
    
    
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub



Private Sub Form_Activate()
    Left = (Screen.width - width) / 2   ' 폼을 가로로 중앙에 놓습니다.
    Top = (Screen.height - height) / 2   ' 폼을 세로로 중앙에 놓습니다.
    
    
    cmb_Model(0).AddItem = "자두이노"
    
    Call Load_Unita_Config
    Call Display_Config
    
    
End Sub

Private Sub Form_Load()
    Call Form_Activate
End Sub



Private Sub Display_Config()

On Error GoTo Err_P:
    If (LANE1_Unita_YN = "Y") Then
        chk_Unita_YN(0).value = 1
    Else
        chk_Unita_YN(0).value = 0
    End If
    txt_UnitaIP(0).text = LANE1_UnitaIP
    txt_UnitaGW(0).text = LANE1_UnitaGW
    txt_UnitaSM(0).text = LANE1_UnitaSM
    txt_UnitaDNS(0).text = LANE1_UnitaDNS
    txt_UnitaLoop(0).text = LANE1_UnitaLoop
    txt_UnitaGate(0).text = LANE1_UnitaGate
    cmb_UnitaModel(0).text = LANE1_UnitaModel
    txt_UnitaActiveTime(0).text = LANE1_UnitaActive_Time
    txt_UnitaRelay_Time(0).text = LANE1_UnitaRelay_Time
    txt_UnitaCPUtemperature(0).text = 0
    
    If (LANE2_Unita_YN = "Y") Then
        chk_Unita_YN(1).value = 1
    Else
        chk_Unita_YN(1).value = 0
    End If
    txt_UnitaIP(1).text = LANE2_UnitaIP
    txt_UnitaGW(1).text = LANE2_UnitaGW
    txt_UnitaSM(1).text = LANE2_UnitaSM
    txt_UnitaDNS(1).text = LANE2_UnitaDNS
    txt_UnitaLoop(1).text = LANE2_UnitaLoop
    txt_UnitaGate(1).text = LANE2_UnitaGate
    cmb_UnitaModel(1).text = LANE2_UnitaModel
    txt_UnitaActiveTime(1).text = LANE2_UnitaActive_Time
    txt_UnitaRelay_Time(1).text = LANE2_UnitaRelay_Time
    txt_UnitaCPUtemperature(1).text = 0
    
    If (LANE3_Unita_YN = "Y") Then
        chk_Unita_YN(2).value = 1
    Else
        chk_Unita_YN(2).value = 0
    End If
    txt_UnitaIP(2).text = LANE3_UnitaIP
    txt_UnitaGW(2).text = LANE3_UnitaGW
    txt_UnitaSM(2).text = LANE3_UnitaSM
    txt_UnitaDNS(2).text = LANE3_UnitaDNS
    txt_UnitaLoop(2).text = LANE3_UnitaLoop
    txt_UnitaGate(2).text = LANE3_UnitaGate
    cmb_UnitaModel(2).text = LANE3_UnitaModel
    txt_UnitaActiveTime(2).text = LANE3_UnitaActive_Time
    txt_UnitaRelay_Time(2).text = LANE3_UnitaRelay_Time
    txt_UnitaCPUtemperature(2).text = 0
    
    If (LANE4_Unita_YN = "Y") Then
        chk_Unita_YN(3).value = 1
    Else
        chk_Unita_YN(3).value = 0
    End If
    txt_UnitaIP(3).text = LANE4_UnitaIP
    txt_UnitaGW(3).text = LANE4_UnitaGW
    txt_UnitaSM(3).text = LANE4_UnitaSM
    txt_UnitaDNS(3).text = LANE4_UnitaDNS
    txt_UnitaLoop(3).text = LANE4_UnitaLoop
    txt_UnitaGate(3).text = LANE4_UnitaGate
    cmb_UnitaModel(3).text = LANE4_UnitaModel
    txt_UnitaActiveTime(3).text = LANE4_UnitaActive_Time
    txt_UnitaRelay_Time(3).text = LANE4_UnitaRelay_Time
    txt_UnitaCPUtemperature(3).text = 0
    
    If (LANE5_Unita_YN = "Y") Then
        chk_Unita_YN(4).value = 1
    Else
        chk_Unita_YN(4).value = 0
    End If
    txt_UnitaIP(4).text = LANE5_UnitaIP
    txt_UnitaGW(4).text = LANE5_UnitaGW
    txt_UnitaSM(4).text = LANE5_UnitaSM
    txt_UnitaDNS(4).text = LANE5_UnitaDNS
    txt_UnitaLoop(4).text = LANE5_UnitaLoop
    txt_UnitaGate(4).text = LANE5_UnitaGate
    cmb_UnitaModel(4).text = LANE5_UnitaModel
    txt_UnitaActiveTime(4).text = LANE5_UnitaActive_Time
    txt_UnitaRelay_Time(4).text = LANE5_UnitaRelay_Time
    txt_UnitaCPUtemperature(4).text = 0
    
    If (LANE6_Unita_YN = "Y") Then
        chk_Unita_YN(5).value = 1
    Else
        chk_Unita_YN(5).value = 0
    End If
    txt_UnitaIP(5).text = LANE6_UnitaIP
    txt_UnitaGW(5).text = LANE6_UnitaGW
    txt_UnitaSM(5).text = LANE6_UnitaSM
    txt_UnitaDNS(5).text = LANE6_UnitaDNS
    txt_UnitaLoop(5).text = LANE6_UnitaLoop
    txt_UnitaGate(5).text = LANE6_UnitaGate
    cmb_UnitaModel(5).text = LANE6_UnitaModel
    txt_UnitaActiveTime(5).text = LANE6_UnitaActive_Time
    txt_UnitaRelay_Time(5).text = LANE6_UnitaRelay_Time
    txt_UnitaCPUtemperature(5).text = 0
    
    Exit Sub
    
Err_P:
    Call DataLogger(" [Unita Display Config err] : " & Err.Description)
End Sub
Private Sub Load_Unita_Config()
    
    

    LANE1_Unita_YN = Get_Ini("System Config", "LANE1_YN", "N")
    LANE1_UnitaIP = Get_Ini("System Config", "LANE1_DeviceIP", "192.168.0.221")
    LANE1_UnitaGW = Get_Ini("System Config", "LANE1_UnitaGW", "192.168.0.1")
    LANE1_UnitaSM = Get_Ini("System Config", "LANE1_UnitaSM", 24)
    LANE1_UnitaDNS = Get_Ini("System Config", "LANE1_UnitaDNS", "168.126.63.1")
    LANE1_UnitaLoop = Get_Ini("System Config", "LANE1_UnitaLoop", 13)
    LANE1_UnitaGate = Get_Ini("System Config", "LANE1_UnitaGate", 26)
    LANE1_UnitaModel = Get_Ini("System Config", "LANE1_UnitaModel", "SampleModel")
    LANE1_UnitaActive_Time = Get_Ini("System Config", "LANE1_UnitaActive_Time", "10")
    LANE1_UnitaRelay_Time = Get_Ini("System Config", "LANE1_UnitaRelay_Time", "0.5")
    'LANE1_UnitaPort = 8888
    
    LANE2_Unita_YN = Get_Ini("System Config", "LANE2_Unita_YN", "N")
    LANE2_UnitaIP = Get_Ini("System Config", "LANE2_UnitaIP", "192.168.0.222")
    LANE2_UnitaGW = Get_Ini("System Config", "LANE2_UnitaGW", "192.168.0.1")
    LANE2_UnitaSM = Get_Ini("System Config", "LANE2_UnitaSM", 24)
    LANE2_UnitaDNS = Get_Ini("System Config", "LANE2_UnitaDNS", "168.126.63.1")
    LANE2_UnitaLoop = Get_Ini("System Config", "LANE2_UnitaLoop", 13)
    LANE2_UnitaGate = Get_Ini("System Config", "LANE2_UnitaGate", 26)
    LANE2_UnitaModel = Get_Ini("System Config", "LANE2_UnitaModel", "SampleModel")
    'LANE2_UnitaPort = 8888
    
    LANE3_Unita_YN = Get_Ini("System Config", "LANE3_Unita_YN", "N")
    LANE3_UnitaIP = Get_Ini("System Config", "LANE3_UnitaIP", "192.168.0.223")
    LANE3_UnitaGW = Get_Ini("System Config", "LANE3_UnitaGW", "192.168.0.1")
    LANE3_UnitaSM = Get_Ini("System Config", "LANE3_UnitaSM", 24)
    LANE3_UnitaDNS = Get_Ini("System Config", "LANE3_UnitaDNS", "168.126.63.1")
    LANE3_UnitaLoop = Get_Ini("System Config", "LANE3_UnitaLoop", 13)
    LANE3_UnitaGate = Get_Ini("System Config", "LANE3_UnitaGate", 26)
    LANE3_UnitaModel = Get_Ini("System Config", "LANE3_UnitaModel", "SampleModel")
    'LANE3_UnitaPort = 8888
    
    LANE4_Unita_YN = Get_Ini("System Config", "LANE4_Unita_YN", "N")
    LANE4_UnitaIP = Get_Ini("System Config", "LANE4_UnitaIP", "192.168.0.224")
    LANE4_UnitaGW = Get_Ini("System Config", "LANE4_UnitaGW", "192.168.0.1")
    LANE4_UnitaSM = Get_Ini("System Config", "LANE4_UnitaSM", 24)
    LANE4_UnitaDNS = Get_Ini("System Config", "LANE4_UnitaDNS", "168.126.63.1")
    LANE4_UnitaLoop = Get_Ini("System Config", "LANE4_UnitaLoop", 13)
    LANE4_UnitaGate = Get_Ini("System Config", "LANE4_UnitaGate", 26)
    LANE4_UnitaModel = Get_Ini("System Config", "LANE4_UnitaModel", "SampleModel")
    'LANE4_UnitaPort = 8888
    
    LANE5_Unita_YN = Get_Ini("System Config", "LANE5_Unita_YN", "N")
    LANE5_UnitaIP = Get_Ini("System Config", "LANE5_UnitaIP", "192.168.0.225")
    LANE5_UnitaGW = Get_Ini("System Config", "LANE5_UnitaGW", "192.168.0.1")
    LANE5_UnitaSM = Get_Ini("System Config", "LANE5_UnitaSM", 24)
    LANE5_UnitaDNS = Get_Ini("System Config", "LANE5_UnitaDNS", "168.126.63.1")
    LANE5_UnitaLoop = Get_Ini("System Config", "LANE5_UnitaLoop", 13)
    LANE5_UnitaGate = Get_Ini("System Config", "LANE5_UnitaGate", 26)
    LANE5_UnitaModel = Get_Ini("System Config", "LANE5_UnitaModel", "SampleModel")
    'LANE5_UnitaPort = 8888
    
    LANE6_Unita_YN = Get_Ini("System Config", "LANE6_Unita_YN", "N")
    LANE6_UnitaIP = Get_Ini("System Config", "LANE6_UnitaIP", "192.168.0.226")
    LANE6_UnitaGW = Get_Ini("System Config", "LANE6_UnitaGW", "192.168.0.1")
    LANE6_UnitaSM = Get_Ini("System Config", "LANE6_UnitaSM", 24)
    LANE6_UnitaDNS = Get_Ini("System Config", "LANE6_UnitaDNS", "168.126.63.1")
    LANE6_UnitaLoop = Get_Ini("System Config", "LANE6_UnitaLoop", 13)
    LANE6_UnitaGate = Get_Ini("System Config", "LANE6_UnitaGate", 26)
    LANE6_UnitaModel = Get_Ini("System Config", "LANE6_UnitaModel", "SampleModel")
    'LANE6_UnitaPort = 8888
    
End Sub

Private Sub Command1_Click()
    

    If (chk_Unita_YN(0).value = 1) Then
        Call Put_Ini("System Config", "LANE1_Unita_YN", "Y")
    Else
        Call Put_Ini("System Config", "LANE1_Unita_YN", "N")
    End If
    Call Put_Ini("System Config", "LANE1_UnitaIP", Trim(txt_UnitaIP(0).text))
    Call Put_Ini("System Config", "LANE1_UnitaGW", Trim(txt_UnitaGW(0).text))
    Call Put_Ini("System Config", "LANE1_UnitaSM", Trim(txt_UnitaSM(0).text))
    Call Put_Ini("System Config", "LANE1_UnitaDNS", Trim(txt_UnitaDNS(0).text))
    Call Put_Ini("System Config", "LANE1_UnitaLoop", Trim(txt_UnitaLoop(0).text))
    Call Put_Ini("System Config", "LANE1_UnitaGate", Trim(txt_UnitaGate(0).text))
    Call Put_Ini("System Config", "LANE1_UnitaModel", UCase(cmb_UnitaModel(0).text))
    Call Put_Ini("System Config", "LANE1_UnitaActive_Time", Trim(txt_UnitaActiveTime(0).text))
    Call Put_Ini("System Config", "LANE1_UnitaRelay_Time", Trim(txt_UnitaRelay_Time(0).text))
        
    If (chk_Unita_YN(1).value = 1) Then
        Call Put_Ini("System Config", "LANE2_Unita_YN", "Y")
    Else
        Call Put_Ini("System Config", "LANE2_Unita_YN", "N")
    End If
    Call Put_Ini("System Config", "LANE2_UnitaIP", Trim(txt_UnitaIP(1).text))
    Call Put_Ini("System Config", "LANE2_UnitaGW", Trim(txt_UnitaGW(1).text))
    Call Put_Ini("System Config", "LANE2_UnitaSM", Trim(txt_UnitaSM(1).text))
    Call Put_Ini("System Config", "LANE2_UnitaDNS", Trim(txt_UnitaDNS(1).text))
    Call Put_Ini("System Config", "LANE2_UnitaLoop", Trim(txt_UnitaLoop(1).text))
    Call Put_Ini("System Config", "LANE2_UnitaGate", Trim(txt_UnitaGate(1).text))
    Call Put_Ini("System Config", "LANE2_UnitaModel", UCase(cmb_UnitaModel(1).text))
    Call Put_Ini("System Config", "LANE2_UnitaActive_Time", Trim(txt_UnitaActiveTime(1).text))
    Call Put_Ini("System Config", "LANE2_UnitaRelay_Time", Trim(txt_UnitaRelay_Time(1).text))
    
    If (chk_Unita_YN(2).value = 1) Then
        Call Put_Ini("System Config", "LANE3_Unita_YN", "Y")
    Else
        Call Put_Ini("System Config", "LANE3_Unita_YN", "N")
    End If
    Call Put_Ini("System Config", "LANE3_UnitaIP", Trim(txt_UnitaIP(2).text))
    Call Put_Ini("System Config", "LANE3_UnitaGW", Trim(txt_UnitaGW(2).text))
    Call Put_Ini("System Config", "LANE3_UnitaSM", Trim(txt_UnitaSM(2).text))
    Call Put_Ini("System Config", "LANE3_UnitaDNS", Trim(txt_UnitaDNS(2).text))
    Call Put_Ini("System Config", "LANE3_UnitaLoop", Trim(txt_UnitaLoop(2).text))
    Call Put_Ini("System Config", "LANE3_UnitaGate", Trim(txt_UnitaGate(2).text))
    Call Put_Ini("System Config", "LANE3_UnitaModel", UCase(cmb_UnitaModel(2).text))
    Call Put_Ini("System Config", "LANE3_UnitaActive_Time", Trim(txt_UnitaActiveTime(2).text))
    Call Put_Ini("System Config", "LANE3_UnitaRelay_Time", Trim(txt_UnitaRelay_Time(2).text))
    
    If (chk_Unita_YN(3).value = 1) Then
        Call Put_Ini("System Config", "LANE4_Unita_YN", "Y")
    Else
        Call Put_Ini("System Config", "LANE4_Unita_YN", "N")
    End If
    Call Put_Ini("System Config", "LANE4_UnitaIP", Trim(txt_UnitaIP(3).text))
    Call Put_Ini("System Config", "LANE4_UnitaGW", Trim(txt_UnitaGW(3).text))
    Call Put_Ini("System Config", "LANE4_UnitaSM", Trim(txt_UnitaSM(3).text))
    Call Put_Ini("System Config", "LANE4_UnitaDNS", Trim(txt_UnitaDNS(3).text))
    Call Put_Ini("System Config", "LANE4_UnitaLoop", Trim(txt_UnitaLoop(3).text))
    Call Put_Ini("System Config", "LANE4_UnitaGate", Trim(txt_UnitaGate(3).text))
    Call Put_Ini("System Config", "LANE4_UnitaModel", UCase(cmb_UnitaModel(3).text))
    Call Put_Ini("System Config", "LANE4_UnitaActive_Time", Trim(txt_UnitaActiveTime(3).text))
    Call Put_Ini("System Config", "LANE4_UnitaRelay_Time", Trim(txt_UnitaRelay_Time(3).text))
    
    If (chk_Unita_YN(4).value = 1) Then
        Call Put_Ini("System Config", "LANE5_Unita_YN", "Y")
    Else
        Call Put_Ini("System Config", "LANE5_Unita_YN", "N")
    End If
    Call Put_Ini("System Config", "LANE5_UnitaIP", Trim(txt_UnitaIP(4).text))
    Call Put_Ini("System Config", "LANE5_UnitaGW", Trim(txt_UnitaGW(4).text))
    Call Put_Ini("System Config", "LANE5_UnitaSM", Trim(txt_UnitaSM(4).text))
    Call Put_Ini("System Config", "LANE5_UnitaDNS", Trim(txt_UnitaDNS(4).text))
    Call Put_Ini("System Config", "LANE5_UnitaLoop", Trim(txt_UnitaLoop(4).text))
    Call Put_Ini("System Config", "LANE5_UnitaGate", Trim(txt_UnitaGate(4).text))
    Call Put_Ini("System Config", "LANE5_UnitaModel", UCase(cmb_UnitaModel(4).text))
    Call Put_Ini("System Config", "LANE5_UnitaActive_Time", Trim(txt_UnitaActiveTime(4).text))
    Call Put_Ini("System Config", "LANE5_UnitaRelay_Time", Trim(txt_UnitaRelay_Time(4).text))
    
    If (chk_Unita_YN(5).value = 1) Then
        Call Put_Ini("System Config", "LANE6_Unita_YN", "Y")
    Else
        Call Put_Ini("System Config", "LANE6_Unita_YN", "N")
    End If
    Call Put_Ini("System Config", "LANE6_UnitaIP", Trim(txt_UnitaIP(5).text))
    Call Put_Ini("System Config", "LANE6_UnitaGW", Trim(txt_UnitaGW(5).text))
    Call Put_Ini("System Config", "LANE6_UnitaSM", Trim(txt_UnitaSM(5).text))
    Call Put_Ini("System Config", "LANE6_UnitaDNS", Trim(txt_UnitaDNS(5).text))
    Call Put_Ini("System Config", "LANE6_UnitaLoop", Trim(txt_UnitaLoop(5).text))
    Call Put_Ini("System Config", "LANE6_UnitaGate", Trim(txt_UnitaGate(5).text))
    Call Put_Ini("System Config", "LANE6_UnitaModel", UCase(cmb_UnitaModel(5).text))
    Call Put_Ini("System Config", "LANE6_UnitaActive_Time", Trim(txt_UnitaActiveTime(5).text))
    Call Put_Ini("System Config", "LANE6_UnitaRelay_Time", Trim(txt_UnitaRelay_Time(5).text))
    
End Sub
