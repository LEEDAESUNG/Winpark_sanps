VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "Threed32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form FrmLPR 
   Caption         =   " ������ ���� ����"
   ClientHeight    =   5490
   ClientLeft      =   9930
   ClientTop       =   6645
   ClientWidth     =   9960
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "FrmLPR.frx":0000
   ScaleHeight     =   5490
   ScaleWidth      =   9960
   Begin MSWinsockLib.Winsock Gate_Winsock 
      Left            =   180
      Top             =   90
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   1110
      Left            =   90
      TabIndex        =   18
      Top             =   4305
      Width           =   9780
   End
   Begin VB.TextBox Txt_Gate 
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   4
      Left            =   1485
      TabIndex        =   4
      Top             =   3855
      Width           =   4485
   End
   Begin VB.TextBox Txt_Gate 
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   3
      Left            =   1485
      TabIndex        =   3
      Top             =   3495
      Width           =   4485
   End
   Begin VB.TextBox Txt_Gate 
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   2
      Left            =   1485
      TabIndex        =   2
      Top             =   3150
      Width           =   1890
   End
   Begin VB.TextBox Txt_Gate 
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   1485
      TabIndex        =   1
      Top             =   2790
      Width           =   1890
   End
   Begin VB.TextBox Txt_Gate 
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   1485
      TabIndex        =   0
      Top             =   2385
      Width           =   1890
   End
   Begin ComctlLib.ListView ListView 
      Height          =   1620
      Left            =   60
      TabIndex        =   12
      Top             =   615
      Width           =   9840
      _ExtentX        =   17357
      _ExtentY        =   2858
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
      Height          =   450
      Index           =   0
      Left            =   8715
      TabIndex        =   9
      Top             =   90
      Width           =   1095
      _Version        =   65536
      _ExtentX        =   1931
      _ExtentY        =   794
      _StockProps     =   78
      Caption         =   "�� ��"
      ForeColor       =   4210752
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
      Height          =   480
      Index           =   4
      Left            =   8850
      TabIndex        =   7
      Top             =   3675
      Width           =   870
      _Version        =   65536
      _ExtentX        =   1535
      _ExtentY        =   847
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
      Picture         =   "FrmLPR.frx":20378
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   480
      Index           =   3
      Left            =   7965
      TabIndex        =   6
      Top             =   3675
      Width           =   870
      _Version        =   65536
      _ExtentX        =   1535
      _ExtentY        =   847
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
      Picture         =   "FrmLPR.frx":206C9
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   480
      Index           =   2
      Left            =   7080
      TabIndex        =   5
      Top             =   3675
      Width           =   870
      _Version        =   65536
      _ExtentX        =   1535
      _ExtentY        =   847
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
      Picture         =   "FrmLPR.frx":20A1A
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   480
      Index           =   1
      Left            =   6195
      TabIndex        =   8
      Top             =   3675
      Width           =   870
      _Version        =   65536
      _ExtentX        =   1535
      _ExtentY        =   847
      _StockProps     =   78
      Caption         =   "�ʱ�ȭ"
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
      Picture         =   "FrmLPR.frx":20D6B
   End
   Begin Threed.SSCommand cmd_Button 
      Height          =   1245
      Index           =   5
      Left            =   6210
      TabIndex        =   10
      Top             =   2325
      Visible         =   0   'False
      Width           =   1725
      _Version        =   65536
      _ExtentX        =   3043
      _ExtentY        =   2196
      _StockProps     =   78
      Caption         =   "OPEN"
      ForeColor       =   32768
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�������"
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
      Index           =   6
      Left            =   7980
      TabIndex        =   11
      Top             =   2325
      Visible         =   0   'False
      Width           =   1725
      _Version        =   65536
      _ExtentX        =   3043
      _ExtentY        =   2196
      _StockProps     =   78
      Caption         =   "CLOSE"
      ForeColor       =   192
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�������"
         Size            =   20.25
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Lbl_Date 
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   3510
      TabIndex        =   19
      Top             =   2430
      Width           =   2325
   End
   Begin VB.Label Lbl_LPR 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��⹮��_02"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   4
      Left            =   225
      TabIndex        =   17
      Top             =   3885
      Width           =   1200
   End
   Begin VB.Label Lbl_LPR 
      BackColor       =   &H00FFFFFF&
      Caption         =   "��⹮��_01"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   3
      Left            =   225
      TabIndex        =   16
      Top             =   3525
      Width           =   1200
   End
   Begin VB.Label Lbl_LPR 
      BackColor       =   &H00FFFFFF&
      Caption         =   "GateIP"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   2
      Left            =   225
      TabIndex        =   15
      Top             =   3165
      Width           =   1200
   End
   Begin VB.Label Lbl_LPR 
      BackColor       =   &H00FFFFFF&
      Caption         =   "GateName"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   1
      Left            =   225
      TabIndex        =   14
      Top             =   2820
      Width           =   1200
   End
   Begin VB.Label Lbl_LPR 
      BackColor       =   &H00FFFFFF&
      Caption         =   "GateNo"
      BeginProperty Font 
         Name            =   "�������"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   225
      TabIndex        =   13
      Top             =   2445
      Width           =   1215
   End
End
Attribute VB_Name = "FrmLPR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tmpDis1, tmpDis2 As String * 32
Dim LprQry, CMD, CMD_IP As String

Private Sub Form_Load()
    Dim i As Integer

    Left = (Screen.Width - Width) / 2   ' ���� ���η� �߾ӿ� �����ϴ�.
    Top = (Screen.Height - Height) / 2   ' ���� ���η� �߾ӿ� �����ϴ�.
    
    '� ���
    
    LprQry = "SELECT * From TB_LPR Order By GateNo"

    Call Clear_Field
    Call ListView_Draw
    Call ListView_SQL
    
    'List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    LPR ����...!!", 0
    Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    �������/���� ����...!!")

End Sub

Public Sub ListView_SQL()
Dim rs As Recordset
Dim Qry As String
Dim itmX As ListItem
Dim INDEX_NO As Long

INDEX_NO = 1
Set rs = New ADODB.Recordset
rs.Open LprQry, adoConn
'lbl_COUNT = rs.RecordCount
Do While Not (rs.EOF)
    Set itmX = ListView.ListItems.Add(, , "" & INDEX_NO)
    itmX.SubItems(1) = "" & rs!GateNo
    itmX.SubItems(2) = "" & rs!GateName
    itmX.SubItems(3) = "" & rs!IP
    itmX.SubItems(4) = "" & rs!Dis1
    itmX.SubItems(5) = "" & rs!Dis2
    itmX.SubItems(6) = "" & rs!RegDate
    rs.MoveNext
    INDEX_NO = INDEX_NO + 1
Loop
Set rs = Nothing
End Sub

Public Sub ListView_Draw()
Dim Column_to_size As Integer

With Me
    Call ListViewExtended(.ListView)
    .ListView.View = lvwReport
    .ListView.ListItems.Clear
    .ListView.ColumnHeaders.Clear
    .ListView.ColumnHeaders.Add , , " No    "
    .ListView.ColumnHeaders.Add , , " GateNo  "
    .ListView.ColumnHeaders.Add , , " GateName      "
    .ListView.ColumnHeaders.Add , , " IP                     "
    .ListView.ColumnHeaders.Add , , " ��⹮��_01                       "
    .ListView.ColumnHeaders.Add , , " ��⹮��_02                       "
    .ListView.ColumnHeaders.Add , , " RegDate                       "
    .ListView.ColumnHeaders.Add , , " "
    
    For Column_to_size = 0 To .ListView.ColumnHeaders.Count - 2
         SendMessage .ListView.hwnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
    Next
End With
End Sub

Private Sub ListView_ItemClick(ByVal Item As ComctlLib.ListItem)
    ListView.SetFocus
    Txt_Gate(0) = ListView.SelectedItem.SubItems(1)
End Sub

Public Sub Clear_Field()
    Dim i As Integer
    
    For i = 0 To 4
        Txt_Gate(i).Text = ""
    Next
    Lbl_Date.Caption = ""
    CMD_IP = ""
End Sub

'������ ����
Sub Delete_Record()
    adoConn.Execute "DELETE FROM TB_LPR WHERE GateNO = '" & Txt_Gate(0) & "'"
    List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & Txt_Gate(0) & "    Gate ���� ���� �Ϸ�", 0
    Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & Txt_Gate(0) & "    Gate ���� ���� �Ϸ�")
    Call ListView_Draw
    Call ListView_SQL
End Sub

Sub Insert_Record()
    Dim rs_COUNT As Recordset
    Dim rs As Recordset
    Dim SQL_COUNT As String
    Dim SQL_QUARY As String
    Dim i As Integer
    Dim Cnt As Integer
    Dim tmp As String
    Dim tmpName, tmpPhone As String
    Dim P As String

    If (Lbl_Date.Caption = "") Then '�űԵ��
        'INSERT
        adoConn.Execute "INSERT INTO TB_LPR VALUES ('" & Txt_Gate(0).Text & "', '" & Txt_Gate(1).Text & "', '" & Txt_Gate(2).Text & "', '" & Txt_Gate(3).Text & "', '" & Txt_Gate(4).Text & "', '" & Format(Now, "YYYYMMDDHHNNSS") & "')"
        'List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_CarNo & "    ������� �Ϸ�", 0
        'Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_CarNo & "    ������� �Ϸ�")
    Else
        adoConn.Execute "UPDATE TB_LPR SET GateNO = '" & Txt_Gate(0).Text & "', GateName = '" & Txt_Gate(1).Text & "', IP = '" & Txt_Gate(2).Text & "', Dis1 = '" & Txt_Gate(3).Text & "', Dis2 = '" & Txt_Gate(4).Text & "' Where RegDate = '" & Lbl_Date.Caption & "'"
        'List1.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_CarNo & "    �������� ���� �Ϸ�", 0
        'Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & txt_CarNo & "    �������� ���� �Ϸ�")
    End If
    
    Call ListView_Draw
    Call ListView_SQL

On Error Resume Next
    If (Err = 3022) Then
        Msg_Box.Label2.Caption = "������ ���̽� ����"
        Msg_Box.Label1.Caption = "�ߺ��� GateNo�� ��������ʽ��ϴ�."
        Msg_Box.Show 1
    End If

End Sub

Private Sub cmd_Button_Click(Index As Integer)
    Dim i, j As Integer
    Dim myExcelFile As New ExcelFile
    Dim tmpFileName As String

    Select Case Index
        Case 0  '����
            Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    ������ ���� ����")
            Unload Me
            Exit Sub
           
        Case 1  '�ʱ�ȭ
            Call Clear_Field
            Exit Sub
        
        Case 2  '��� & ����
            If (Lbl_Date.Caption = "") Then
                If (Data_Error_Check = False) Then
                    Msg_Box.Label2.Caption = "�ʵ� �Է� ����"
                    Msg_Box.Label1.Caption = "�߿��� �׸��� �Է����� �ʾҽ��ϴ�."
                    Msg_Box.Show 1
                Else
                    Call Insert_Record
                    CMD = ""
                    CMD = "CMD_Display"
                    Call Socket_ConnectGate(CMD_IP, 233)
                    Call Clear_Field
                End If
            Else
                Msg_Box.Label2.Caption = "�ű� ������ �Է� ����"
                Msg_Box.Label1.Caption = "�ű� �����Ͱ� �ƴմϴ�." & vbCrLf & vbCrLf & " �ٽ� �ѹ� Ȯ���ϼ���."
                Msg_Box.Show 1
                Call Clear_Field
            End If
            Exit Sub
            
        Case 3   'Update
            If (Lbl_Date.Caption <> "") Then
                If (Data_Error_Check = True) Then
                    Call Insert_Record
                    CMD = ""
                    CMD = "CMD_Display"
                    Call Socket_ConnectGate(CMD_IP, 233)
                    Call Clear_Field
                    Exit Sub
                End If
            End If
            
                
        Case 4  'Delete
            If (Lbl_Date.Caption = "") Then
               Exit Sub
            End If
            Call Delete_Record
            Call Clear_Field
            Exit Sub
    
        Case 5
            If (Len(CMD_IP) <> 0) Then
                CMD = ""
                CMD = "CMD_RELAY_01"
                i = MsgBox(CMD_IP & "   " & "���ܱ� ����", vbYesNo)
                Select Case i
                    Case Is = vbYes
                        Call Socket_ConnectGate(CMD_IP, 233)
                    Case Is = vbNo
                End Select
            End If
            Exit Sub
        Case 6
            If (Len(CMD_IP) <> 0) Then
                CMD = ""
                CMD = "CMD_RELAY_02"
                i = MsgBox(CMD_IP & "   " & "���ܱ� �ݱ�", vbYesNo)
                Select Case i
                    Case Is = vbYes
                        Call Socket_ConnectGate(CMD_IP, 233)
                    Case Is = vbNo
                End Select
            End If
            Exit Sub
    End Select

On Error Resume Next

End Sub

Private Sub sOutput(ByVal strText As String, ByVal strIP As String)
    List1.AddItem Format(Now, "hh:nn:ss") & "     " & strText & "     " & strIP, 0
End Sub

Private Sub Socket_ConnectGate(ByVal IP As String, ByVal Port As Long)
    'Gate_Winsock.Close

    If (Gate_Winsock.State <> sckClosed) Then
        Gate_Winsock.Close
        'DoEvents
    End If
    Gate_Winsock.Connect IP, Port

    Call sOutput("[Gate ����]", CMD & "  " & IP)
    'Call Err_doc("    [Gate ����]  �õ� IP = " & IP & "    PORT = " & Port)
End Sub

Private Sub Gate_Winsock_Connect()
    Dim bdata() As Byte

    ReDim bdata(Len(CMD) - 1) As Byte
    bdata = StrConv(CMD, vbFromUnicode)
    Gate_Winsock.SendData bdata

    Call sOutput("[Gate �۽�]", CMD)
'    If (Check5.value = 1) Then
'        Call Err_doc("    [Gate �۽�] " & CMD)
'    End If
    'Fee_sock.Close
End Sub

Private Sub Gate_Winsock_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String

    Gate_Winsock.GetData strData, , bytesTotal
    Call sOutput("[Gate ����]", strData)

    Gate_Winsock.Close
End Sub

Private Sub Gate_Winsock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    Call sOutput(Source, "[Gate ����] " & "���� : " & Description)
'    If (Check5.value = 1) Then
'        Call Err_doc("   [Gate ����] " & "���� : " & Description)
'    End If
End Sub

'�ʼ� �Է� ������ Ȯ��

Private Function Data_Error_Check()
    Dim Error_Flag As Boolean
        
    Error_Flag = True
    
    If (IsNumeric(Txt_Gate(0).Text) = False) Then
        Txt_Gate(0).Text = "���ڸ� �Է��ϼ���...!!"
        Error_Flag = False
    End If
    
    If (LenH(Txt_Gate(2).Text) = 0) Then
        Txt_Gate(0).Text = "IP �Է��ϼ���...!!"
        Error_Flag = False
    End If
    
    Data_Error_Check = Error_Flag

End Function

Private Sub Txt_Gate_Change(Index As Integer)

    Select Case Index
        Case 0
            If (LenH(Txt_Gate(0)) <> 0) Then
                Call Search_Record
            End If
        Case Else
    
    End Select
End Sub

Sub Search_Record()
    Dim rs As Recordset
    Dim SQL_SEARCH As String
    Dim itmX As ListItem
    Dim INDEX_NO As Long
    
    SQL_SEARCH = "SELECT * From TB_LPR WHERE GateNO = '" & Txt_Gate(0) & "'"
    'Debug.Print SQL_SEARCH
    Set rs = New ADODB.Recordset
    rs.Open SQL_SEARCH, adoConn
    
    If (rs.RecordCount <> 0) Then
        Txt_Gate(0).Text = "" & rs!GateNo
        Txt_Gate(1).Text = "" & rs!GateName
        Txt_Gate(2).Text = "" & rs!IP
        Txt_Gate(3).Text = "" & rs!Dis1
        Txt_Gate(4).Text = "" & rs!Dis2
        Lbl_Date.Caption = "" & rs!RegDate
        CMD_IP = ""
        CMD_IP = "" & rs!IP
    End If
    Set rs = Nothing

End Sub

'����Ű �Է½� �� ����
'���Ӽ� keypreview = true ����
Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    SendKeys "{TAB}"
End If

End Sub






