VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMctl32.OCX"
Begin VB.Form FormIPCamera 
   BackColor       =   &H00404040&
   BorderStyle     =   1  '´ÜÀÏ °íÁ¤
   Caption         =   "ParkingManager¢â"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12405
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   12405
   StartUpPosition =   3  'Windows ±âº»°ª
   Begin VB.TextBox txt_CCTV_EX2 
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Text            =   "¸Þ¸ð¶õ"
      Top             =   6855
      Width           =   2880
   End
   Begin VB.TextBox txt_CCTV_EX1 
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Text            =   "Á¤¹®ÀÔ±¸1"
      Top             =   6075
      Width           =   2880
   End
   Begin VB.TextBox txt_CCTV_COMMENT 
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Text            =   "³»ºÎ¸Á01"
      Top             =   5295
      Width           =   2880
   End
   Begin VB.TextBox txt_CCTV_URL 
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Text            =   "rtsp://jawootek.iptime.org:11554/user=admin&password=&channel=1&stream=0.sdp"
      Top             =   4560
      Width           =   10395
   End
   Begin ComctlLib.ListView ListView_CCTV 
      Height          =   3000
      Left            =   135
      TabIndex        =   0
      Top             =   1185
      Width           =   12075
      _ExtentX        =   21299
      _ExtentY        =   5292
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   16777215
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ"
         Size            =   9.75
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin Threed.SSCommand SSCommand2 
      Cancel          =   -1  'True
      Height          =   570
      Left            =   10935
      TabIndex        =   2
      Top             =   135
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   1005
      _StockProps     =   78
      Caption         =   "´Ý±â"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "FormIPCamera.frx":0000
   End
   Begin Threed.SSCommand SSCommand3 
      Height          =   690
      Left            =   1800
      TabIndex        =   11
      Top             =   7800
      Width           =   1665
      _Version        =   65536
      _ExtentX        =   2937
      _ExtentY        =   1217
      _StockProps     =   78
      Caption         =   "ÃÊ±âÈ­"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "FormIPCamera.frx":0351
   End
   Begin Threed.SSCommand SSCommand4 
      Height          =   690
      Left            =   3480
      TabIndex        =   12
      Top             =   7800
      Width           =   1665
      _Version        =   65536
      _ExtentX        =   2937
      _ExtentY        =   1217
      _StockProps     =   78
      Caption         =   "µî·Ï"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "FormIPCamera.frx":06A2
   End
   Begin Threed.SSCommand SSCommand5 
      Height          =   690
      Left            =   5160
      TabIndex        =   13
      Top             =   7800
      Width           =   1665
      _Version        =   65536
      _ExtentX        =   2937
      _ExtentY        =   1217
      _StockProps     =   78
      Caption         =   "¼öÁ¤"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "FormIPCamera.frx":09F3
   End
   Begin Threed.SSCommand SSCommand6 
      Height          =   690
      Left            =   6840
      TabIndex        =   14
      Top             =   7800
      Width           =   1665
      _Version        =   65536
      _ExtentX        =   2937
      _ExtentY        =   1217
      _StockProps     =   78
      Caption         =   "»èÁ¦"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "³ª´®°íµñ"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   1
      RoundedCorners  =   0   'False
      Picture         =   "FormIPCamera.frx":0D44
   End
   Begin VB.Label lbl_ex2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Åõ¸í
      Caption         =   "¸Þ¸ð"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   285
      TabIndex        =   10
      Top             =   6960
      Width           =   540
   End
   Begin VB.Label lbl_ex1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Åõ¸í
      Caption         =   "¼³Ä¡Àå¼Ò"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   285
      TabIndex        =   9
      Top             =   6180
      Width           =   1080
   End
   Begin VB.Label lbl_comment 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Åõ¸í
      Caption         =   "¸ÁÁ¾·ù"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   285
      TabIndex        =   8
      Top             =   5430
      Width           =   810
   End
   Begin VB.Label lbl_cctv_url 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Åõ¸í
      Caption         =   "Á¢¼ÓÁÖ¼Ò"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   14.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   285
      TabIndex        =   7
      Top             =   4635
      Width           =   1080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Index           =   0
      X1              =   135
      X2              =   12225
      Y1              =   780
      Y2              =   795
   End
   Begin VB.Label lbl_APS 
      BackStyle       =   0  'Åõ¸í
      Caption         =   " CCTV ¼³Á¤"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
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
      Left            =   75
      TabIndex        =   1
      Top             =   300
      Width           =   4185
   End
End
Attribute VB_Name = "FormIPCamera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Selected_CCTV_URL As String

Private Sub Form_Load()

    Left = (Screen.width - width) / 2   ' ÆûÀ» °¡·Î·Î Áß¾Ó¿¡ ³õ½À´Ï´Ù.
    Top = (Screen.height - height) / 2   ' ÆûÀ» ¼¼·Î·Î Áß¾Ó¿¡ ³õ½À´Ï´Ù.
    
    Call Clear_Field
    Call ListView_CCTV_Draw
    Call ListView_CCTV_SQL
    
End Sub


Private Sub SSCommand2_Click()
    Unload Me
End Sub

Private Sub Clear_Field()
    txt_CCTV_URL = ""
    txt_CCTV_COMMENT = ""
    txt_CCTV_EX1 = ""
    txt_CCTV_EX2 = ""
End Sub

Private Sub ListView_CCTV_SQL()
    Dim bQryResult As Boolean
    Dim rs As Recordset
    Dim qry As String
    Dim itmX As ListItem
    Dim INDEX_NO As Long
    Dim i As Integer
    
    
    Set rs = New ADODB.Recordset
    'rs.Open RegQry, adoConn
    qry = "SELECT * From tb_CCTV"
    bQryResult = DataBaseQuery(rs, adoConn, qry, False)
    If (bQryResult = False) Then
        Exit Sub
    End If

    INDEX_NO = 1
    Do While Not (rs.EOF)
        Set itmX = ListView_CCTV.ListItems.Add(, , "" & INDEX_NO)
        
        i = 1
        itmX.SubItems(i) = "" & rs!url: i = i + 1
        itmX.SubItems(i) = "" & rs!Comments: i = i + 1
        itmX.SubItems(i) = "" & rs!EX1: i = i + 1
        itmX.SubItems(i) = "" & rs!EX2: i = i + 1

        rs.MoveNext
        INDEX_NO = INDEX_NO + 1
    Loop
    Set rs = Nothing
End Sub

Private Sub ListView_CCTV_Draw()
    Dim Column_to_size As Integer
    
    With Me
        Call ListViewExtended(.ListView_CCTV)
        .ListView_CCTV.View = lvwReport
        .ListView_CCTV.ListItems.Clear
        .ListView_CCTV.ColumnHeaders.Clear
        .ListView_CCTV.ColumnHeaders.Add , , " No   "
        .ListView_CCTV.ColumnHeaders.Add , , " Á¢¼ÓÁÖ¼Ò                                                                                                                  "
        .ListView_CCTV.ColumnHeaders.Add , , " ¸Á Á¾·ù     "
        .ListView_CCTV.ColumnHeaders.Add , , " ±¸ºÐ    "
        .ListView_CCTV.ColumnHeaders.Add , , " ¸Þ¸ð              "
        
        For Column_to_size = 0 To .ListView_CCTV.ColumnHeaders.Count - 2
             SendMessage .ListView_CCTV.hwnd, LVM_SETCOLUMNWIDTH, Column_to_size, LVSCW_AUTOSIZE_USEHEADER
        Next
    End With

End Sub



Private Sub ListView_CCTV_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
    Dim i As Integer
    With ListView_CCTV
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

Private Sub ListView_CCTV_ItemClick(ByVal Item As ComctlLib.ListItem)
    On Error Resume Next
    
    ListView_CCTV.SetFocus
    
    txt_CCTV_URL = ListView_CCTV.SelectedItem.SubItems(1)
    txt_CCTV_COMMENT = ListView_CCTV.SelectedItem.SubItems(2)
    txt_CCTV_EX1 = ListView_CCTV.SelectedItem.SubItems(3)
    txt_CCTV_EX2 = ListView_CCTV.SelectedItem.SubItems(4)

    Selected_CCTV_URL = txt_CCTV_URL
End Sub

'ÃÊ±âÈ­
Private Sub SSCommand3_Click()
    Call Clear_Field
End Sub

'ÀÔ·Â
Private Sub SSCommand4_Click()
    If (Len(txt_CCTV_URL) <= 0) Then
        Exit Sub
    End If
    
    adoConn.Execute "INSERT INTO tb_CCTV (URL, COMMENTS, EX1, EX2) VALUES ('" & txt_CCTV_URL & "', '" & txt_CCTV_COMMENT & "', '" & txt_CCTV_EX1 & "', '" & txt_CCTV_EX2 & "')"
    adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('" & txt_CCTV_EX1 & "', 'HOST', 'CCTV RTSPÁÖ¼Ò µî·Ï', '', " & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
    
    Call Clear_Field
    Call ListView_CCTV_Draw
    Call ListView_CCTV_SQL
End Sub

'¼öÁ¤
Private Sub SSCommand5_Click()
    If (Len(Selected_CCTV_URL) <= 0 Or Len(txt_CCTV_URL) <= 0) Then
        Exit Sub
    End If
    
    adoConn.Execute "UPDATE tb_CCTV SET URL = '" & txt_CCTV_URL & "', COMMENTS = '" & txt_CCTV_COMMENT & "', EX1 = '" & txt_CCTV_EX1 & "', EX2 = '" & txt_CCTV_EX2 & "' WHERE URL = '" & Selected_CCTV_URL & "' "
    adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('" & txt_CCTV_EX1 & "', 'HOST', 'CCTV RTSPÁÖ¼Ò ¼öÁ¤', '', " & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
    
    Selected_CCTV_URL = ""
    Call Clear_Field
    Call ListView_CCTV_Draw
    Call ListView_CCTV_SQL
End Sub

'»èÁ¦
Private Sub SSCommand6_Click()
    If (Len(Selected_CCTV_URL) <= 0 Or Len(txt_CCTV_URL) <= 0) Then
        Exit Sub
    End If
    
    adoConn.Execute "DELETE FROM tb_CCTV WHERE URL = '" & Selected_CCTV_URL & "' "
    adoConn.Execute "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('" & txt_CCTV_EX1 & "', 'HOST', 'CCTV RTSPÁÖ¼Ò »èÁ¦', '', " & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
    Selected_CCTV_URL = ""
    
    Call Clear_Field
    Call ListView_CCTV_Draw
    Call ListView_CCTV_SQL
End Sub


