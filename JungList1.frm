VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form JungList1 
   Appearance      =   0  '���
   BackColor       =   &H80000005&
   Caption         =   " ����� ����"
   ClientHeight    =   14925
   ClientLeft      =   26685
   ClientTop       =   705
   ClientWidth     =   19140
   LinkTopic       =   "Form1"
   Picture         =   "JungList1.frx":0000
   ScaleHeight     =   14925
   ScaleWidth      =   19140
   Begin Threed.SSCommand Command2 
      Height          =   780
      Left            =   17220
      TabIndex        =   4
      Top             =   13320
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   1376
      _StockProps     =   78
      Caption         =   "Exit"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
      Picture         =   "JungList1.frx":F651
   End
   Begin Threed.SSCommand Command1 
      Height          =   780
      Left            =   15810
      TabIndex        =   3
      Top             =   13320
      Visible         =   0   'False
      Width           =   1275
      _Version        =   65536
      _ExtentX        =   2249
      _ExtentY        =   1376
      _StockProps     =   78
      Caption         =   "�μ�"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      RoundedCorners  =   0   'False
      Picture         =   "JungList1.frx":F9A2
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "JungList1.frx":FCF3
      Left            =   10410
      List            =   "JungList1.frx":FD06
      Style           =   2  '��Ӵٿ� ���
      TabIndex        =   2
      Top             =   13500
      Width           =   2115
   End
   Begin Crystal.CrystalReport Report1 
      Bindings        =   "JungList1.frx":FD32
      Left            =   19380
      Top             =   675
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "C:\Winpark\Data\jlist.rpt"
      WindowLeft      =   0
      WindowTop       =   0
      WindowWidth     =   1024
      WindowHeight    =   768
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
   End
   Begin MSDataGridLib.DataGrid DBGrid1 
      Bindings        =   "JungList1.frx":FD46
      Height          =   10260
      Left            =   375
      TabIndex        =   1
      Top             =   1620
      Width           =   18450
      _ExtentX        =   32544
      _ExtentY        =   18098
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      ColumnHeaders   =   -1  'True
      HeadLines       =   2
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   11.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9.75
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "������ȣ"
         Caption         =   " ������ȣ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "����"
         Caption         =   " ������"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "�̸�"
         Caption         =   " �̸�"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "����"
         Caption         =   " �� ��"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "��ȭ��ȣ"
         Caption         =   " ��ȭ��ȣ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "�߱���"
         Caption         =   " �߱���"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "�߱޽ð�"
         Caption         =   " Upadate"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "������"
         Caption         =   " ������"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "������"
         Caption         =   " ������"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "���"
         Caption         =   " �� ��"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1042
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         ScrollBars      =   2
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   1695.118
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1500.095
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1260.284
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2085.166
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1649.764
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   3000.189
         EndProperty
         BeginProperty Column07 
         EndProperty
         BeginProperty Column08 
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   3825.071
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   19350
      Top             =   165
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "jawootek"
      OtherAttributes =   ""
      UserName        =   "root"
      Password        =   "root"
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   2  '��� ����
      BackColor       =   &H00404040&
      Caption         =   "���� ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   8250
      TabIndex        =   5
      Top             =   13500
      Width           =   1710
   End
   Begin VB.Label LblRecordCount 
      Alignment       =   2  '��� ����
      BackColor       =   &H00000000&
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   11115
      TabIndex        =   0
      Top             =   1005
      Width           =   3315
   End
End
Attribute VB_Name = "JungList1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Order$

Private Sub Combo1_Click()
Order$ = " ORDER BY " & Combo1.List(Combo1.ListIndex)
'Record_Source = "SELECT * FROM regcar WHERE ������ >= '" & Format(Now, "yyyy-mm-dd") & "'" & Order$
Record_Source = "SELECT * FROM regcar " & Order$
Adodc1.RecordSource = Record_Source
Adodc1.Refresh
End Sub

Private Sub Form_Activate()
If (Adodc1.Recordset.RecordCount <> 0) Then
    Adodc1.Recordset.MoveLast
End If
LblRecordCount.Caption = "������� : " & Adodc1.Recordset.RecordCount & " ��"
End Sub

Private Sub Form_Load()
Left = (Screen.Width - Width) / 2   ' ���� ���η� �߾ӿ� �����ϴ�.
Top = (Screen.Height - Height) / 2   ' ���� ���η� �߾ӿ� �����ϴ�.
'Adodc1.ConnectionString = AdoConn_Str
Report1.Connect = AdoConn_Str
'Adodc1.RecordSource = "SELECT * FROM regcar WHERE ������ >= '" & Format(Now, "yyyy-mm-dd") & "'"
Adodc1.RecordSource = "SELECT * FROM regcar"
Adodc1.Refresh
Combo1.ListIndex = 0
End Sub


Private Sub Command1_Click()
 Dim tmp%
 Dim SelectionFormula$
 On Error GoTo PrintReportError
 'SelectionFormula$ = "{regcar.������} >= '" & Format(Now, "yyyy-mm-dd") & "'"
 DoEvents
 Report1.ReportFileName = Report_Path_Name$ & "jlist.rpt"
 
 Select Case Combo1.ListIndex
        Case 0
             Report1.SortFields(0) = "+{regcar.������ȣ}"
        Case 1
             Report1.SortFields(0) = "+{regcar.�̸�}"
        Case 2
             '����
             'Report1.SortFields(0) = "+{regcar.�Ҽ�}"
             Report1.SortFields(0) = "+{regcar.����}"
        Case 3
             '����
             'Report1.SortFields(0) = "+{regcar.�Ҽ�}"
             Report1.SortFields(0) = "+{regcar.������}"
 End Select
 
 
 Report1.SelectionFormula = SelectionFormula$
 Report1.CopiesToPrinter = 1
 Report1.Action = 1
 Exit Sub
PrintReportError:
Msg_Box.Label2.Caption = "����Ʈ �۾� ����"
Msg_Box.Label1.Caption = "�����Ͱ� �����۵� �����ʽ��ϴ�."
Msg_Box.Show 1
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

