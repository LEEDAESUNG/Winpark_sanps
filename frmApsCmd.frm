VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmApsCmd 
   BorderStyle     =   1  '´ÜÀÏ °íÁ¤
   Caption         =   "APS Command"
   ClientHeight    =   9045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5070
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   5070
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3300
      Top             =   90
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1110
      Left            =   90
      TabIndex        =   28
      Top             =   7890
      Width           =   3960
   End
   Begin VB.CommandButton btn_ApsCmd 
      Caption         =   "ÃÊ±â È­¸é"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   9
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   2430
      TabIndex        =   25
      Top             =   3330
      Width           =   2055
   End
   Begin VB.TextBox txt_APScmd 
      Alignment       =   2  '°¡¿îµ¥ ¸ÂÃã
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5610
      TabIndex        =   20
      Text            =   "1"
      Top             =   2430
      Width           =   975
   End
   Begin MSWinsockLib.Winsock CMD_Sock 
      Left            =   3870
      Top             =   60
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmd_Exit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   9
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4140
      TabIndex        =   8
      Top             =   7950
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "APS Commad "
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   9
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5805
      Left            =   90
      TabIndex        =   7
      Top             =   2010
      Width           =   4905
      Begin VB.CommandButton btn_ApsCmd 
         Caption         =   "¿µ ¼ö Áõ"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   9
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   2340
         TabIndex        =   24
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton btn_ApsCmd 
         Caption         =   "ÁöÁ¤ ½Ã°¢"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   9
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   150
         TabIndex        =   21
         Top             =   4650
         Width           =   2055
      End
      Begin VB.TextBox txt_FeeCmd 
         Alignment       =   1  '¿À¸¥ÂÊ ¸ÂÃã
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2370
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "0"
         Top             =   4200
         Width           =   1455
      End
      Begin VB.CommandButton btn_ApsCmd 
         Caption         =   "ÁöÁ¤ ±Ý¾×"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   9
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   150
         TabIndex        =   17
         Top             =   4170
         Width           =   2055
      End
      Begin VB.CommandButton btn_ApsCmd 
         Caption         =   "Â÷´Ü±â ¿­¸²"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   9
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   7
         Left            =   150
         TabIndex        =   16
         Top             =   3690
         Width           =   2055
      End
      Begin VB.CommandButton btn_ApsCmd 
         Caption         =   "100¿ø ÇÒÀÎ"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   9
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   150
         TabIndex        =   15
         Top             =   3210
         Width           =   2055
      End
      Begin VB.CommandButton btn_ApsCmd 
         Caption         =   "4½Ã°£ ÇÒÀÎ"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   9
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   150
         TabIndex        =   14
         Top             =   2730
         Width           =   2055
      End
      Begin VB.CommandButton btn_ApsCmd 
         Caption         =   "2½Ã°£ ÇÒÀÎ"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   9
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   150
         TabIndex        =   13
         Top             =   2250
         Width           =   2055
      End
      Begin VB.CommandButton btn_ApsCmd 
         Caption         =   "1½Ã°£ ÇÒÀÎ"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   9
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   150
         TabIndex        =   12
         Top             =   1770
         Width           =   2055
      End
      Begin VB.CommandButton btn_ApsCmd 
         Caption         =   "Á¤»ê Ãë¼Ò"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   9
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   150
         TabIndex        =   11
         Top             =   1320
         Width           =   2055
      End
      Begin VB.CommandButton btn_ApsCmd 
         Caption         =   "Àü¾× ÇÒÀÎ"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   9
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   150
         TabIndex        =   10
         Top             =   840
         Width           =   2055
      End
      Begin VB.CommandButton btn_ApsCmd 
         Caption         =   "ÀÏÀÏÁÖÂ÷¿ä±Ý"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   9
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   150
         TabIndex        =   9
         Top             =   360
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   345
         Left            =   2370
         TabIndex        =   22
         Top             =   4650
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
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
         Format          =   111411200
         CurrentDate     =   36927
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   345
         Left            =   2370
         TabIndex        =   23
         Top             =   5100
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "³ª´®°íµñ"
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
         Format          =   111411202
         CurrentDate     =   36927
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " APS Info. "
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   9
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1755
      Left            =   90
      TabIndex        =   0
      Top             =   180
      Width           =   4905
      Begin VB.Label lbl_ApsPort 
         Caption         =   "5888"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1410
         TabIndex        =   27
         Top             =   1260
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "APS Port :"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   240
         TabIndex        =   26
         Top             =   1260
         Width           =   1215
      End
      Begin VB.Label lbl_CmdPort 
         Caption         =   "5888"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1410
         TabIndex        =   6
         Top             =   930
         Width           =   1215
      End
      Begin VB.Label lbl_ApsIP 
         Caption         =   "192.168.0.9"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1410
         TabIndex        =   5
         Top             =   630
         Width           =   2625
      End
      Begin VB.Label lbl_ApsName 
         Caption         =   "1001"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1410
         TabIndex        =   4
         Top             =   330
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "CMD Port :"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   930
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "APS IP :"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   630
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "APS Name :"
         BeginProperty Font 
            Name            =   "³ª´®°íµñ"
            Size            =   9.75
            Charset         =   129
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   330
         Width           =   1215
      End
   End
   Begin VB.Label lbl_Rcv 
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   9
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   19
      Top             =   6960
      Width           =   3045
   End
End
Attribute VB_Name = "frmApsCmd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_ApsCmd_Click(Index As Integer)

For i = 0 To 11
    btn_ApsCmd(i).Enabled = False
    Timer1.Enabled = True
Next i

'20150922010101
' 2 + 20

Select Case Index
    Case 0  'ÀÏÀÏÁÖÂ÷¿ä±Ý
        Glo_APSCMD_Str = "00"
    Case 1  'Àü¾×ÇÒÀÎ
        Glo_APSCMD_Str = "01100"
    Case 2  'Á¤»êÃë¼Ò
        Glo_APSCMD_Str = "10"   'Cancel
    Case 3  '1½Ã°£ ÇÒÀÎ
        Glo_APSCMD_Str = "0260"
    Case 4  '2½Ã°£ ÇÒÀÎ
        Glo_APSCMD_Str = "02120"
    Case 5  '4½Ã°£ ÇÒÀÎ
        Glo_APSCMD_Str = "02240"
    Case 6  '100¿ø ÇÒÀÎ
        Glo_APSCMD_Str = "03100"
    Case 7  'Â÷´Ü±â ¿­¸²
        Glo_APSCMD_Str = "12"
    Case 8  'ÁöÁ¤±Ý¾×
        'Glo_APSCMD_Str = "CMD_WON  " & " & txt_FeeCmd & "
    Case 9  'ÁöÁ¤½Ã°£
        Glo_APSCMD_Str = "04" & Format(DTPicker1.value, "yyyymmdd") & Format(DTPicker2.value, "hhnnss")
    Case 10 '¿µ¼öÁõ ¹ßÇà
        Glo_APSCMD_Str = "13"
    Case 11 'È­¸é ÃÊ±âÈ­
        Glo_APSCMD_Str = "11"
End Select

Call CMD_Connect

End Sub

Private Sub CMD_Connect()
Dim bData() As Byte

On Error GoTo Err_P

    If (CMD_Sock.State <> sckClosed) Then
        CMD_Sock.Close
    End If
    
    Select Case Left(Glo_APSCMD_Str, 2)
        Case "04", "13" 'ÃÊ±âÈ­¸é 5889
            CMD_Sock.Connect Glo_Aps_IP, Glo_Aps_PORT
        Case Else   '°è»êÈ­¸é 5888
            CMD_Sock.Connect Glo_Aps_IP, Glo_APSCMD_Port
    End Select
    
Exit Sub

Err_P:
    Call DataLogger("[CMD_Connect] Err_Msg : " & Err.Description)

End Sub

Private Sub cmd_Exit_Click()
    Call DataLogger("[APS CMD Button] : APS CMD Á¾·á")
    Unload Me
    Exit Sub
End Sub

Private Sub CMD_Sock_Connect()
Dim sdata As String
Dim bData() As Byte
Dim i As Integer

On Error GoTo Err_P

sdata = Glo_APSCMD_Str
ReDim bData(Len(sdata) - 1) As Byte
bData = StrConv(sdata, vbFromUnicode)
CMD_Sock.SendData bData
Call DataLogger("[APS CMD SND]  SND : " & Glo_APSCMD_Str)
List1.AddItem Format(Now, "HH:NN:SS") & " SND : " & Glo_APSCMD_Str, 0
Glo_APSCMD_Str = ""

Exit Sub

Err_P:
    Call DataLogger(" [CMDSock_Connect Proc] Err_Msg : " & Err.Description)

End Sub

Private Sub CMD_Sock_DataArrival(ByVal bytesTotal As Long)
Dim rMsg As String
Dim B() As Byte
Dim Ret As Integer
Dim i As Integer
Dim sdata As String

On Error GoTo Err_P

ReDim B(bytesTotal - 1)

CMD_Sock.GetData B(), vbArray + vbByte, bytesTotal
For i = 0 To bytesTotal - 1
    If (B(i) >= &H80) Then
        rMsg = rMsg & Chr$(Val("&H" & Hex(B(i)) & Hex(B(i + 1))))
        i = i + 1
    Else
        rMsg = rMsg & Chr$(B(i))
    End If
Next i

Call DataLogger("[APS CMD RCV]  RCV : " & rMsg)
List1.AddItem Format(Now, "HH:NN:SS") & " RCV : " & rMsg, 0

CMD_Sock.Close

Exit Sub

Err_P:
    Call DataLogger(" [APS CMD RCV] Err_Msg : " & Err.Description)

End Sub

Private Sub Form_Load()

    Left = (Screen.width - width) / 2   ' ÆûÀ» °¡·Î·Î Áß¾Ó¿¡ ³õ½À´Ï´Ù.
    Top = (Screen.height - height) / 2   ' ÆûÀ» ¼¼·Î·Î Áß¾Ó¿¡ ³õ½À´Ï´Ù.
    
    lbl_ApsIP.Caption = Glo_Aps_IP
    lbl_CmdPort.Caption = Glo_APSCMD_Port
    lbl_ApsPort.Caption = Glo_Aps_PORT
    DTPicker1.value = Now
    DTPicker2.value = Now
    
    List1.Clear
    CMD_Sock.Protocol = sckTCPProtocol
    
    
End Sub

'Æû¼Ó¼º keypreview = true ¼³Á¤
Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyEscape Then
        KeyAscii = 0
        Unload Me
    End If

End Sub

Private Sub Timer1_Timer()

For i = 0 To 11
    btn_ApsCmd(i).Enabled = True
    Timer1.Enabled = False
Next i

btn_ApsCmd(8).Enabled = False

End Sub
