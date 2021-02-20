VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.UserControl Connect 
   ClientHeight    =   945
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2355
   ScaleHeight     =   945
   ScaleWidth      =   2355
   Begin Threed.SSPanel Frame11 
      Height          =   825
      Left            =   45
      TabIndex        =   0
      Top             =   60
      Width           =   2265
      _Version        =   65536
      _ExtentX        =   3995
      _ExtentY        =   1455
      _StockProps     =   15
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "굴림"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin Threed.SSPanel Led1 
         Height          =   135
         Index           =   0
         Left            =   1770
         TabIndex        =   1
         Top             =   420
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   238
         _StockProps     =   15
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin Threed.SSPanel Led1 
         Height          =   135
         Index           =   1
         Left            =   1770
         TabIndex        =   2
         Top             =   585
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   238
         _StockProps     =   15
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "굴림"
            Size            =   12
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         RoundedCorners  =   0   'False
         Outline         =   -1  'True
      End
      Begin VB.Image Image1 
         Height          =   420
         Left            =   90
         Picture         =   "Connect.ctx":0000
         Top             =   375
         Width           =   1620
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "OFF Line"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   195
         TabIndex        =   3
         Top             =   120
         Width           =   1395
      End
      Begin VB.Image Image2 
         Height          =   420
         Left            =   90
         Picture         =   "Connect.ctx":06A2
         Top             =   375
         Width           =   1620
      End
   End
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'기본 속성 값:
Const m_def_OnLine_Chk = 0
Const m_def_ForeColor = 0
Const m_def_TxColor = 0
Const m_def_RxColor = 0
Const m_def_Enabled = 0
Const m_def_Status = 0
Const m_def_OnLine_Flag = 0

'속성 변수:
Public On_Line_Cnt As Integer
Dim m_Font As Font
Dim m_RxColor As Long
Dim m_TxColor As Long
Dim m_Status As Boolean
Dim m_OnLine_Flag As Boolean
Dim m_ProcessNum As String

Public Property Get ProcessNum() As String
    ProcessNum = m_ProcessNum
End Property

Public Property Let ProcessNum(ByVal New_ProcessNum As String)
    m_ProcessNum = New_ProcessNum
    PropertyChanged "ProcessNum"
End Property

Public Property Get OnLine_Flag() As Boolean
    OnLine_Flag = m_OnLine_Flag
End Property

Public Property Let OnLine_Flag(ByVal New_Flag As Boolean)
    m_OnLine_Flag = New_Flag
    PropertyChanged "OnLine_Flag"
End Property

Public Property Get Caption() As String
    Caption = Frame11.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    Frame11.Caption = New_Caption
    PropertyChanged "Caption"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = Frame11.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Frame11.ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get Status() As Boolean
    Status = m_Status
End Property

Public Property Let Status(ByVal New_Status As Boolean)
    m_Status = New_Status
    If (m_Status = True) Then
        Label1.Caption = "On Line...."
        Label1.ForeColor = vbBlack
        Image1.Visible = False
        Image2.Visible = True
    Else
        Label1.Caption = "Off Line...."
        Label1.ForeColor = vbRed
        Image1.Visible = True
        Image2.Visible = False
    End If
    PropertyChanged "Status"
End Property

Public Property Get RxColor() As OLE_COLOR
Attribute RxColor.VB_Description = "개체에서 텍스트나 그래픽을 표시하는 전경색을 반환하거나 설정합니다."
    RxColor = m_RxColor
    'RxColor = Led1(1).BackColor
End Property

Public Property Let RxColor(ByVal New_RxColor As OLE_COLOR)
    m_RxColor = New_RxColor
    Led1(1).BackColor = New_RxColor
    PropertyChanged "RxColor"
End Property

Public Property Get TxColor() As OLE_COLOR
    TxColor = m_TxColor
    'TxColor = Led1(0).BackColor
End Property

Public Property Let TxColor(ByVal New_TxColor As OLE_COLOR)
    m_TxColor = New_TxColor
    Led1(0).BackColor = New_TxColor
    PropertyChanged "TxColor"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Font 개체를 반환합니다."
Attribute Font.VB_UserMemId = -512
    Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set m_Font = New_Font
    Frame11.Font = m_Font
    PropertyChanged "Font"
End Property

'사용자 정의 컨트롤 속성을 초기화합니다.
Private Sub UserControl_InitProperties()
Frame11.ForeColor = m_def_ForeColor
Frame11.Caption = "Connect"
Led1(0).BackColor = &H808080
Led1(1).BackColor = &H808080
Label1.Caption = "Off Line...."
Label1.ForeColor = vbRed
m_Status = m_def_Status
m_OnLine_Flag = m_def_OnLine_Flag
Image1.Visible = True
Image2.Visible = False
End Sub

'속성을 읽는다....
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Frame11.ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    Frame11.Caption = PropBag.ReadProperty("Caption", "")
    'Led1(0).BackColor = PropBag.ReadProperty("TxColor", m_def_TxColor)
    'Led1(1).BackColor = PropBag.ReadProperty("RxColor", m_def_RxColor)
    
    m_TxColor = PropBag.ReadProperty("TxColor", m_def_TxColor)
    m_RxColor = PropBag.ReadProperty("RxColor", m_def_RxColor)
    Led1(0).BackColor = m_TxColor
    Led1(1).BackColor = m_RxColor
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_Status = PropBag.ReadProperty("Status", m_def_Status)
    If (m_Status = True) Then
        Label1.Caption = "On Line...."
        Label1.ForeColor = vbBlack
        Image1.Visible = False
        Image2.Visible = True
    Else
        Label1.Caption = "Off Line...."
        Label1.ForeColor = vbRed
        Image1.Visible = True
        Image2.Visible = False
    End If
End Sub

'속성을 씁니다.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", Frame11.Caption, "")
    Call PropBag.WriteProperty("ForeColor", Frame11.ForeColor, m_def_ForeColor)
    'Call PropBag.WriteProperty("TxColor", Led1(0).BackColor, m_def_TxColor)
    'Call PropBag.WriteProperty("RxColor", Led1(1).BackColor, m_def_RxColor)
    Call PropBag.WriteProperty("TxColor", m_TxColor, m_def_TxColor)
    Call PropBag.WriteProperty("RxColor", m_RxColor, m_def_RxColor)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("Status", m_Status, m_def_Status)
End Sub

Public Property Get OnLine_Chk() As Boolean
    Dim PauseTime As Single
    Dim start  As Single
    Dim Ret As Boolean
    PauseTime = 0.5
    start = Timer
    Me.OnLine_Flag = False
    Me.TxColor = vbRed
    Do While Timer < start + PauseTime
        DoEvents
        If (Timer < start) Then
            start = start - 86400
        End If
        If (Me.OnLine_Flag = True) Then
            Me.TxColor = &HC0C0C0
            Me.RxColor = &HC0C0C0
            Ret = True
            Me.Status = True
            On_Line_Cnt = 0
            Exit Do
        End If
    Loop
    Me.TxColor = &HC0C0C0
    Me.RxColor = &HC0C0C0
    If (On_Line_Cnt >= 1) Then
        Me.Status = False
        OnLine_Chk = False
        Me.OnLine_Flag = False
    Else
        On_Line_Cnt = On_Line_Cnt + 1
        OnLine_Chk = Me.OnLine_Flag
    End If
End Property

