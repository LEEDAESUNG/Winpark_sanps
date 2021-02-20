VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FormSound 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  '단일 고정
   Caption         =   "ParkingManager™"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   19290
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   19290
   StartUpPosition =   3  'Windows 기본값
   Begin VB.Frame Frame9 
      BackColor       =   &H00E0E0E0&
      Height          =   3525
      Index           =   0
      Left            =   -30
      TabIndex        =   165
      Top             =   930
      Width           =   6360
      Begin VB.CheckBox chk_SND_GuestRegCarExpDate 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   0
         Left            =   2865
         TabIndex        =   324
         Top             =   3165
         Width           =   1200
      End
      Begin VB.CommandButton btn_SND_GuestRegCarExpDate 
         Caption         =   "..."
         Height          =   300
         Index           =   0
         Left            =   2235
         TabIndex        =   323
         Top             =   3120
         Width           =   300
      End
      Begin VB.TextBox txt_Str_GuestRegCarExpDate 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   4245
         TabIndex        =   322
         Text            =   "Text1"
         Top             =   3120
         Width           =   1020
      End
      Begin VB.CommandButton btn_PLY_GuestRegCarExpDate 
         Caption         =   "▶"
         Height          =   300
         Index           =   0
         Left            =   2550
         TabIndex        =   321
         Top             =   3120
         Width           =   255
      End
      Begin VB.ComboBox cmb_Disp1EmgColorGuestRegCarExpDate 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         ItemData        =   "FormSound.frx":0000
         Left            =   5280
         List            =   "FormSound.frx":0019
         Style           =   2  '드롭다운 목록
         TabIndex        =   320
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   3120
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorGuestRegCarExpDate 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         ItemData        =   "FormSound.frx":0039
         Left            =   5790
         List            =   "FormSound.frx":0052
         Style           =   2  '드롭다운 목록
         TabIndex        =   319
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   3120
         Width           =   525
      End
      Begin VB.CheckBox chk_SND_GuestRegCar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   0
         Left            =   2865
         TabIndex        =   318
         Top             =   2850
         Width           =   1200
      End
      Begin VB.CommandButton btn_SND_GuestRegCar 
         Caption         =   "..."
         Height          =   300
         Index           =   0
         Left            =   2235
         TabIndex        =   317
         Top             =   2805
         Width           =   300
      End
      Begin VB.TextBox txt_Str_GuestRegCar 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   4245
         TabIndex        =   316
         Text            =   "Text1"
         Top             =   2805
         Width           =   1020
      End
      Begin VB.CommandButton btn_PLY_GuestRegCar 
         Caption         =   "▶"
         Height          =   300
         Index           =   0
         Left            =   2550
         TabIndex        =   315
         Top             =   2805
         Width           =   255
      End
      Begin VB.ComboBox cmb_Disp1EmgColorGuestRegCar 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         ItemData        =   "FormSound.frx":0072
         Left            =   5280
         List            =   "FormSound.frx":008B
         Style           =   2  '드롭다운 목록
         TabIndex        =   314
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   2805
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorGuestRegCar 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         ItemData        =   "FormSound.frx":00AB
         Left            =   5790
         List            =   "FormSound.frx":00C4
         Style           =   2  '드롭다운 목록
         TabIndex        =   313
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   2805
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorRegExpDate 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         ItemData        =   "FormSound.frx":00E4
         Left            =   5790
         List            =   "FormSound.frx":00FD
         Style           =   2  '드롭다운 목록
         TabIndex        =   240
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   2490
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp1EmgColorRegExpDate 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         ItemData        =   "FormSound.frx":011D
         Left            =   5280
         List            =   "FormSound.frx":0136
         Style           =   2  '드롭다운 목록
         TabIndex        =   239
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   2490
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorDay 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         ItemData        =   "FormSound.frx":0156
         Left            =   5790
         List            =   "FormSound.frx":016F
         Style           =   2  '드롭다운 목록
         TabIndex        =   238
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   2175
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp1EmgColorDay 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         ItemData        =   "FormSound.frx":018F
         Left            =   5280
         List            =   "FormSound.frx":01A8
         Style           =   2  '드롭다운 목록
         TabIndex        =   237
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   2175
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorTaxi 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         ItemData        =   "FormSound.frx":01C8
         Left            =   5790
         List            =   "FormSound.frx":01E1
         Style           =   2  '드롭다운 목록
         TabIndex        =   236
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   1845
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp1EmgColorTaxi 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         ItemData        =   "FormSound.frx":0201
         Left            =   5280
         List            =   "FormSound.frx":021A
         Style           =   2  '드롭다운 목록
         TabIndex        =   235
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   1845
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorBKList 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         ItemData        =   "FormSound.frx":023A
         Left            =   5790
         List            =   "FormSound.frx":0253
         Style           =   2  '드롭다운 목록
         TabIndex        =   234
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   1530
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp1EmgColorBKList 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         ItemData        =   "FormSound.frx":0273
         Left            =   5280
         List            =   "FormSound.frx":028C
         Style           =   2  '드롭다운 목록
         TabIndex        =   233
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   1530
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorNoRec 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         ItemData        =   "FormSound.frx":02AC
         Left            =   5790
         List            =   "FormSound.frx":02C5
         Style           =   2  '드롭다운 목록
         TabIndex        =   232
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   1215
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp1EmgColorNoRec 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         ItemData        =   "FormSound.frx":02E5
         Left            =   5280
         List            =   "FormSound.frx":02FE
         Style           =   2  '드롭다운 목록
         TabIndex        =   231
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   1215
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorGuest 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         ItemData        =   "FormSound.frx":031E
         Left            =   5790
         List            =   "FormSound.frx":0337
         Style           =   2  '드롭다운 목록
         TabIndex        =   230
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   900
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp1EmgColorGuest 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         ItemData        =   "FormSound.frx":0357
         Left            =   5280
         List            =   "FormSound.frx":0370
         Style           =   2  '드롭다운 목록
         TabIndex        =   229
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   900
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorReg 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         ItemData        =   "FormSound.frx":0390
         Left            =   5790
         List            =   "FormSound.frx":03A9
         Style           =   2  '드롭다운 목록
         TabIndex        =   228
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   585
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp1EmgColorReg 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         ItemData        =   "FormSound.frx":03C9
         Left            =   5280
         List            =   "FormSound.frx":03E2
         Style           =   2  '드롭다운 목록
         TabIndex        =   227
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   585
         Width           =   525
      End
      Begin VB.CheckBox chk_SND_RegExpDate 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   0
         Left            =   2865
         TabIndex        =   201
         Top             =   2535
         Width           =   1200
      End
      Begin VB.CommandButton btn_PLY_RegExpDate 
         Caption         =   "▶"
         Height          =   300
         Index           =   0
         Left            =   2550
         TabIndex        =   200
         Top             =   2490
         Width           =   255
      End
      Begin VB.TextBox txt_Str_RegExpDate 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   4245
         TabIndex        =   199
         Text            =   "Text1"
         Top             =   2490
         Width           =   1020
      End
      Begin VB.CommandButton btn_SND_RegExpDate 
         Caption         =   "..."
         Height          =   300
         Index           =   0
         Left            =   2235
         TabIndex        =   198
         Top             =   2490
         Width           =   300
      End
      Begin VB.CommandButton btn_SND_Day 
         Caption         =   "..."
         Height          =   300
         Index           =   0
         Left            =   2235
         TabIndex        =   189
         Top             =   2175
         Width           =   300
      End
      Begin VB.CheckBox chk_SND_Day 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   0
         Left            =   2865
         TabIndex        =   188
         Top             =   2220
         Width           =   1200
      End
      Begin VB.TextBox txt_Str_Day 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   4245
         TabIndex        =   187
         Text            =   "Text1"
         Top             =   2175
         Width           =   1020
      End
      Begin VB.CommandButton btn_PLY_Day 
         Caption         =   "▶"
         Height          =   300
         Index           =   0
         Left            =   2550
         TabIndex        =   186
         Top             =   2175
         Width           =   255
      End
      Begin VB.CommandButton btn_PLY_Taxi 
         Caption         =   "▶"
         Height          =   300
         Index           =   0
         Left            =   2550
         TabIndex        =   185
         Top             =   1845
         Width           =   255
      End
      Begin VB.TextBox txt_Str_Taxi 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   4245
         TabIndex        =   184
         Text            =   "Text1"
         Top             =   1845
         Width           =   1020
      End
      Begin VB.CheckBox chk_SND_Taxi 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   0
         Left            =   2865
         TabIndex        =   183
         Top             =   1890
         Width           =   1200
      End
      Begin VB.CommandButton btn_SND_Taxi 
         Caption         =   "..."
         Height          =   300
         Index           =   0
         Left            =   2235
         TabIndex        =   182
         Top             =   1845
         Width           =   300
      End
      Begin VB.CheckBox chk_SND_NoRec 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   0
         Left            =   2865
         TabIndex        =   181
         Top             =   1260
         Width           =   1200
      End
      Begin VB.CheckBox chk_SND_Guest 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   0
         Left            =   2865
         TabIndex        =   180
         Top             =   945
         Width           =   1200
      End
      Begin VB.CheckBox chk_SND_BKList 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   0
         Left            =   2865
         TabIndex        =   179
         Top             =   1575
         Width           =   1200
      End
      Begin VB.TextBox txt_Str_Guest 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   4245
         TabIndex        =   178
         Text            =   "Text1"
         Top             =   900
         Width           =   1020
      End
      Begin VB.TextBox txt_Str_NoRec 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   4245
         TabIndex        =   177
         Text            =   "Text1"
         Top             =   1215
         Width           =   1020
      End
      Begin VB.TextBox txt_Str_BKList 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   4245
         TabIndex        =   176
         Text            =   "Text1"
         Top             =   1530
         Width           =   1020
      End
      Begin VB.TextBox txt_Str_Reg 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   4245
         TabIndex        =   175
         Text            =   "Text1"
         Top             =   585
         Width           =   1020
      End
      Begin VB.CommandButton btn_SND_NoRec 
         Caption         =   "..."
         Height          =   300
         Index           =   0
         Left            =   2235
         TabIndex        =   174
         Top             =   1215
         Width           =   300
      End
      Begin VB.CommandButton btn_SND_Guest 
         Caption         =   "..."
         Height          =   300
         Index           =   0
         Left            =   2235
         TabIndex        =   173
         Top             =   900
         Width           =   300
      End
      Begin VB.CommandButton btn_SND_BKList 
         Caption         =   "..."
         Height          =   300
         Index           =   0
         Left            =   2235
         TabIndex        =   172
         Top             =   1530
         Width           =   300
      End
      Begin VB.CommandButton btn_PLY_NoRec 
         Caption         =   "▶"
         Height          =   300
         Index           =   0
         Left            =   2550
         TabIndex        =   171
         Top             =   1215
         Width           =   255
      End
      Begin VB.CommandButton btn_PLY_Guest 
         Caption         =   "▶"
         Height          =   300
         Index           =   0
         Left            =   2550
         TabIndex        =   170
         Top             =   900
         Width           =   255
      End
      Begin VB.CommandButton btn_PLY_BKList 
         Caption         =   "▶"
         Height          =   300
         Index           =   0
         Left            =   2550
         TabIndex        =   169
         Top             =   1530
         Width           =   255
      End
      Begin VB.CommandButton btn_PLY_Reg 
         Caption         =   "▶"
         Height          =   300
         Index           =   0
         Left            =   2550
         TabIndex        =   168
         Top             =   585
         Width           =   255
      End
      Begin VB.CommandButton btn_SND_Reg 
         Caption         =   "..."
         Height          =   300
         Index           =   0
         Left            =   2235
         TabIndex        =   167
         Top             =   585
         Width           =   300
      End
      Begin VB.CheckBox chk_SND_Reg 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   0
         Left            =   2865
         TabIndex        =   166
         Top             =   630
         Width           =   1200
      End
      Begin VB.Label lbl_GuestRegCarExpDate 
         BackColor       =   &H00E0E0E0&
         Caption         =   "방문예약만료"
         Height          =   210
         Index           =   0
         Left            =   420
         TabIndex        =   312
         Top             =   3165
         Width           =   1335
      End
      Begin VB.Label lbl_GuestRegCar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "방문예약차량"
         Height          =   210
         Index           =   0
         Left            =   420
         TabIndex        =   311
         Top             =   2850
         Width           =   1335
      End
      Begin VB.Label lbl_RegExpDate 
         BackColor       =   &H00E0E0E0&
         Caption         =   "기간만료차량"
         Height          =   270
         Index           =   0
         Left            =   420
         TabIndex        =   197
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label21 
         BackColor       =   &H00E0E0E0&
         Caption         =   "요일제위반차량"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   420
         TabIndex        =   196
         Top             =   2190
         Width           =   1680
      End
      Begin VB.Label Label21 
         BackColor       =   &H00E0E0E0&
         Caption         =   "영업차량"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   420
         TabIndex        =   195
         Top             =   1860
         Width           =   1680
      End
      Begin VB.Label Label20 
         BackColor       =   &H00E0E0E0&
         Caption         =   "미인식차량"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   420
         TabIndex        =   194
         Top             =   1230
         Width           =   1680
      End
      Begin VB.Label Label19 
         BackColor       =   &H00E0E0E0&
         Caption         =   "미등록차량"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   420
         TabIndex        =   193
         Top             =   915
         Width           =   1680
      End
      Begin VB.Label Label18 
         BackColor       =   &H00E0E0E0&
         Caption         =   "출입제한차량"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   420
         TabIndex        =   192
         Top             =   1545
         Width           =   1680
      End
      Begin VB.Label Label19 
         BackColor       =   &H00E0E0E0&
         Caption         =   "등록차량"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   420
         TabIndex        =   191
         Top             =   600
         Width           =   1680
      End
      Begin VB.Label lbl_Lane 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00E0E0E0&
         Caption         =   "Lane1"
         BeginProperty Font 
            Name            =   "나눔고딕 ExtraBold"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   45
         TabIndex        =   190
         Top             =   165
         Width           =   675
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00E0E0E0&
      Height          =   3525
      Index           =   1
      Left            =   6465
      TabIndex        =   133
      Top             =   930
      Width           =   6360
      Begin VB.CheckBox chk_SND_GuestRegCarExpDate 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   1
         Left            =   2865
         TabIndex        =   338
         Top             =   3180
         Width           =   1200
      End
      Begin VB.CommandButton btn_SND_GuestRegCarExpDate 
         Caption         =   "..."
         Height          =   300
         Index           =   1
         Left            =   2235
         TabIndex        =   337
         Top             =   3135
         Width           =   300
      End
      Begin VB.TextBox txt_Str_GuestRegCarExpDate 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   4245
         TabIndex        =   336
         Text            =   "Text1"
         Top             =   3135
         Width           =   1020
      End
      Begin VB.CommandButton btn_PLY_GuestRegCarExpDate 
         Caption         =   "▶"
         Height          =   300
         Index           =   1
         Left            =   2550
         TabIndex        =   335
         Top             =   3135
         Width           =   255
      End
      Begin VB.ComboBox cmb_Disp1EmgColorGuestRegCarExpDate 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         ItemData        =   "FormSound.frx":0402
         Left            =   5280
         List            =   "FormSound.frx":041B
         Style           =   2  '드롭다운 목록
         TabIndex        =   334
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   3135
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorGuestRegCarExpDate 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         ItemData        =   "FormSound.frx":043B
         Left            =   5790
         List            =   "FormSound.frx":0454
         Style           =   2  '드롭다운 목록
         TabIndex        =   333
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   3135
         Width           =   525
      End
      Begin VB.CheckBox chk_SND_GuestRegCar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   1
         Left            =   2865
         TabIndex        =   332
         Top             =   2865
         Width           =   1200
      End
      Begin VB.CommandButton btn_SND_GuestRegCar 
         Caption         =   "..."
         Height          =   300
         Index           =   1
         Left            =   2235
         TabIndex        =   331
         Top             =   2820
         Width           =   300
      End
      Begin VB.TextBox txt_Str_GuestRegCar 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   4245
         TabIndex        =   330
         Text            =   "Text1"
         Top             =   2820
         Width           =   1020
      End
      Begin VB.CommandButton btn_PLY_GuestRegCar 
         Caption         =   "▶"
         Height          =   300
         Index           =   1
         Left            =   2550
         TabIndex        =   329
         Top             =   2820
         Width           =   255
      End
      Begin VB.ComboBox cmb_Disp1EmgColorGuestRegCar 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         ItemData        =   "FormSound.frx":0474
         Left            =   5280
         List            =   "FormSound.frx":048D
         Style           =   2  '드롭다운 목록
         TabIndex        =   328
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   2820
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorGuestRegCar 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         ItemData        =   "FormSound.frx":04AD
         Left            =   5790
         List            =   "FormSound.frx":04C6
         Style           =   2  '드롭다운 목록
         TabIndex        =   327
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   2820
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp1EmgColorReg 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         ItemData        =   "FormSound.frx":04E6
         Left            =   5280
         List            =   "FormSound.frx":04FF
         Style           =   2  '드롭다운 목록
         TabIndex        =   254
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   585
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorReg 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         ItemData        =   "FormSound.frx":051F
         Left            =   5790
         List            =   "FormSound.frx":0538
         Style           =   2  '드롭다운 목록
         TabIndex        =   253
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   585
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp1EmgColorGuest 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         ItemData        =   "FormSound.frx":0558
         Left            =   5280
         List            =   "FormSound.frx":0571
         Style           =   2  '드롭다운 목록
         TabIndex        =   252
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   900
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorGuest 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         ItemData        =   "FormSound.frx":0591
         Left            =   5790
         List            =   "FormSound.frx":05AA
         Style           =   2  '드롭다운 목록
         TabIndex        =   251
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   900
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp1EmgColorNoRec 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         ItemData        =   "FormSound.frx":05CA
         Left            =   5280
         List            =   "FormSound.frx":05E3
         Style           =   2  '드롭다운 목록
         TabIndex        =   250
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   1215
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorNoRec 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         ItemData        =   "FormSound.frx":0603
         Left            =   5790
         List            =   "FormSound.frx":061C
         Style           =   2  '드롭다운 목록
         TabIndex        =   249
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   1215
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp1EmgColorBKList 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         ItemData        =   "FormSound.frx":063C
         Left            =   5280
         List            =   "FormSound.frx":0655
         Style           =   2  '드롭다운 목록
         TabIndex        =   248
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   1530
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorBKList 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         ItemData        =   "FormSound.frx":0675
         Left            =   5790
         List            =   "FormSound.frx":068E
         Style           =   2  '드롭다운 목록
         TabIndex        =   247
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   1530
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp1EmgColorTaxi 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         ItemData        =   "FormSound.frx":06AE
         Left            =   5280
         List            =   "FormSound.frx":06C7
         Style           =   2  '드롭다운 목록
         TabIndex        =   246
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   1845
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorTaxi 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         ItemData        =   "FormSound.frx":06E7
         Left            =   5790
         List            =   "FormSound.frx":0700
         Style           =   2  '드롭다운 목록
         TabIndex        =   245
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   1845
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp1EmgColorDay 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         ItemData        =   "FormSound.frx":0720
         Left            =   5280
         List            =   "FormSound.frx":0739
         Style           =   2  '드롭다운 목록
         TabIndex        =   244
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   2175
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorDay 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         ItemData        =   "FormSound.frx":0759
         Left            =   5790
         List            =   "FormSound.frx":0772
         Style           =   2  '드롭다운 목록
         TabIndex        =   243
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   2175
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp1EmgColorRegExpDate 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         ItemData        =   "FormSound.frx":0792
         Left            =   5280
         List            =   "FormSound.frx":07AB
         Style           =   2  '드롭다운 목록
         TabIndex        =   242
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   2490
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorRegExpDate 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         ItemData        =   "FormSound.frx":07CB
         Left            =   5790
         List            =   "FormSound.frx":07E4
         Style           =   2  '드롭다운 목록
         TabIndex        =   241
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   2490
         Width           =   525
      End
      Begin VB.CheckBox chk_SND_RegExpDate 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   1
         Left            =   2865
         TabIndex        =   206
         Top             =   2550
         Width           =   1200
      End
      Begin VB.CommandButton btn_PLY_RegExpDate 
         Caption         =   "▶"
         Height          =   300
         Index           =   1
         Left            =   2550
         TabIndex        =   205
         Top             =   2505
         Width           =   255
      End
      Begin VB.TextBox txt_Str_RegExpDate 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   4245
         TabIndex        =   204
         Text            =   "Text1"
         Top             =   2505
         Width           =   1020
      End
      Begin VB.CommandButton btn_SND_RegExpDate 
         Caption         =   "..."
         Height          =   300
         Index           =   1
         Left            =   2235
         TabIndex        =   203
         Top             =   2505
         Width           =   300
      End
      Begin VB.CheckBox chk_SND_Reg 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   1
         Left            =   2865
         TabIndex        =   157
         Top             =   630
         Width           =   1200
      End
      Begin VB.CommandButton btn_SND_Reg 
         Caption         =   "..."
         Height          =   300
         Index           =   1
         Left            =   2235
         TabIndex        =   156
         Top             =   585
         Width           =   300
      End
      Begin VB.CommandButton btn_PLY_Reg 
         Caption         =   "▶"
         Height          =   300
         Index           =   1
         Left            =   2550
         TabIndex        =   155
         Top             =   585
         Width           =   255
      End
      Begin VB.CommandButton btn_PLY_BKList 
         Caption         =   "▶"
         Height          =   300
         Index           =   1
         Left            =   2550
         TabIndex        =   154
         Top             =   1530
         Width           =   255
      End
      Begin VB.CommandButton btn_PLY_Guest 
         Caption         =   "▶"
         Height          =   300
         Index           =   1
         Left            =   2550
         TabIndex        =   153
         Top             =   900
         Width           =   255
      End
      Begin VB.CommandButton btn_PLY_NoRec 
         Caption         =   "▶"
         Height          =   300
         Index           =   1
         Left            =   2550
         TabIndex        =   152
         Top             =   1215
         Width           =   255
      End
      Begin VB.CommandButton btn_SND_BKList 
         Caption         =   "..."
         Height          =   300
         Index           =   1
         Left            =   2235
         TabIndex        =   151
         Top             =   1530
         Width           =   300
      End
      Begin VB.CommandButton btn_SND_Guest 
         Caption         =   "..."
         Height          =   300
         Index           =   1
         Left            =   2235
         TabIndex        =   150
         Top             =   900
         Width           =   300
      End
      Begin VB.CommandButton btn_SND_NoRec 
         Caption         =   "..."
         Height          =   300
         Index           =   1
         Left            =   2235
         TabIndex        =   149
         Top             =   1215
         Width           =   300
      End
      Begin VB.TextBox txt_Str_Reg 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   4245
         TabIndex        =   148
         Text            =   "Text1"
         Top             =   585
         Width           =   1020
      End
      Begin VB.TextBox txt_Str_BKList 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   4245
         TabIndex        =   147
         Text            =   "Text1"
         Top             =   1530
         Width           =   1020
      End
      Begin VB.TextBox txt_Str_NoRec 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   4245
         TabIndex        =   146
         Text            =   "Text1"
         Top             =   1215
         Width           =   1020
      End
      Begin VB.TextBox txt_Str_Guest 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   4245
         TabIndex        =   145
         Text            =   "Text1"
         Top             =   900
         Width           =   1020
      End
      Begin VB.CheckBox chk_SND_BKList 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   1
         Left            =   2865
         TabIndex        =   144
         Top             =   1575
         Width           =   1200
      End
      Begin VB.CheckBox chk_SND_Guest 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   1
         Left            =   2865
         TabIndex        =   143
         Top             =   945
         Width           =   1200
      End
      Begin VB.CheckBox chk_SND_NoRec 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   1
         Left            =   2865
         TabIndex        =   142
         Top             =   1260
         Width           =   1200
      End
      Begin VB.CommandButton btn_SND_Taxi 
         Caption         =   "..."
         Height          =   300
         Index           =   1
         Left            =   2235
         TabIndex        =   141
         Top             =   1845
         Width           =   300
      End
      Begin VB.CheckBox chk_SND_Taxi 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   1
         Left            =   2865
         TabIndex        =   140
         Top             =   1890
         Width           =   1200
      End
      Begin VB.TextBox txt_Str_Taxi 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   4245
         TabIndex        =   139
         Text            =   "Text1"
         Top             =   1845
         Width           =   1020
      End
      Begin VB.CommandButton btn_PLY_Taxi 
         Caption         =   "▶"
         Height          =   300
         Index           =   1
         Left            =   2550
         TabIndex        =   138
         Top             =   1845
         Width           =   255
      End
      Begin VB.CommandButton btn_PLY_Day 
         Caption         =   "▶"
         Height          =   300
         Index           =   1
         Left            =   2550
         TabIndex        =   137
         Top             =   2175
         Width           =   255
      End
      Begin VB.TextBox txt_Str_Day 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   4245
         TabIndex        =   136
         Text            =   "Text1"
         Top             =   2175
         Width           =   1020
      End
      Begin VB.CheckBox chk_SND_Day 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   1
         Left            =   2865
         TabIndex        =   135
         Top             =   2220
         Width           =   1200
      End
      Begin VB.CommandButton btn_SND_Day 
         Caption         =   "..."
         Height          =   300
         Index           =   1
         Left            =   2235
         TabIndex        =   134
         Top             =   2175
         Width           =   300
      End
      Begin VB.Label lbl_GuestRegCarExpDate 
         BackColor       =   &H00E0E0E0&
         Caption         =   "방문예약만료"
         Height          =   210
         Index           =   1
         Left            =   420
         TabIndex        =   326
         Top             =   3180
         Width           =   1335
      End
      Begin VB.Label lbl_GuestRegCar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "방문예약차량"
         Height          =   210
         Index           =   1
         Left            =   420
         TabIndex        =   325
         Top             =   2865
         Width           =   1335
      End
      Begin VB.Label lbl_RegExpDate 
         BackColor       =   &H00E0E0E0&
         Caption         =   "기간만료차량"
         Height          =   270
         Index           =   1
         Left            =   420
         TabIndex        =   202
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label lbl_Lane 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00E0E0E0&
         Caption         =   "Lane2"
         BeginProperty Font 
            Name            =   "나눔고딕 ExtraBold"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   45
         TabIndex        =   164
         Top             =   165
         Width           =   675
      End
      Begin VB.Label Label19 
         BackColor       =   &H00E0E0E0&
         Caption         =   "등록차량"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   420
         TabIndex        =   163
         Top             =   600
         Width           =   1680
      End
      Begin VB.Label Label18 
         BackColor       =   &H00E0E0E0&
         Caption         =   "출입제한차량"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   420
         TabIndex        =   162
         Top             =   1545
         Width           =   1680
      End
      Begin VB.Label Label19 
         BackColor       =   &H00E0E0E0&
         Caption         =   "미등록차량"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   420
         TabIndex        =   161
         Top             =   915
         Width           =   1680
      End
      Begin VB.Label Label20 
         BackColor       =   &H00E0E0E0&
         Caption         =   "미인식차량"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   420
         TabIndex        =   160
         Top             =   1230
         Width           =   1680
      End
      Begin VB.Label Label21 
         BackColor       =   &H00E0E0E0&
         Caption         =   "영업차량"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   420
         TabIndex        =   159
         Top             =   1860
         Width           =   1680
      End
      Begin VB.Label Label21 
         BackColor       =   &H00E0E0E0&
         Caption         =   "요일제위반차량"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   420
         TabIndex        =   158
         Top             =   2190
         Width           =   1680
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00E0E0E0&
      Height          =   3525
      Index           =   2
      Left            =   12960
      TabIndex        =   101
      Top             =   930
      Width           =   6360
      Begin VB.CheckBox chk_SND_GuestRegCarExpDate 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   2
         Left            =   2865
         TabIndex        =   352
         Top             =   3180
         Width           =   1200
      End
      Begin VB.CommandButton btn_SND_GuestRegCarExpDate 
         Caption         =   "..."
         Height          =   300
         Index           =   2
         Left            =   2235
         TabIndex        =   351
         Top             =   3135
         Width           =   300
      End
      Begin VB.TextBox txt_Str_GuestRegCarExpDate 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   4245
         TabIndex        =   350
         Text            =   "Text1"
         Top             =   3135
         Width           =   1020
      End
      Begin VB.CommandButton btn_PLY_GuestRegCarExpDate 
         Caption         =   "▶"
         Height          =   300
         Index           =   2
         Left            =   2550
         TabIndex        =   349
         Top             =   3135
         Width           =   255
      End
      Begin VB.ComboBox cmb_Disp1EmgColorGuestRegCarExpDate 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         ItemData        =   "FormSound.frx":0804
         Left            =   5280
         List            =   "FormSound.frx":081D
         Style           =   2  '드롭다운 목록
         TabIndex        =   348
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   3135
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorGuestRegCarExpDate 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         ItemData        =   "FormSound.frx":083D
         Left            =   5790
         List            =   "FormSound.frx":0856
         Style           =   2  '드롭다운 목록
         TabIndex        =   347
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   3135
         Width           =   525
      End
      Begin VB.CheckBox chk_SND_GuestRegCar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   2
         Left            =   2865
         TabIndex        =   346
         Top             =   2865
         Width           =   1200
      End
      Begin VB.CommandButton btn_SND_GuestRegCar 
         Caption         =   "..."
         Height          =   300
         Index           =   2
         Left            =   2235
         TabIndex        =   345
         Top             =   2820
         Width           =   300
      End
      Begin VB.TextBox txt_Str_GuestRegCar 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   4245
         TabIndex        =   344
         Text            =   "Text1"
         Top             =   2820
         Width           =   1020
      End
      Begin VB.CommandButton btn_PLY_GuestRegCar 
         Caption         =   "▶"
         Height          =   300
         Index           =   2
         Left            =   2550
         TabIndex        =   343
         Top             =   2820
         Width           =   255
      End
      Begin VB.ComboBox cmb_Disp1EmgColorGuestRegCar 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         ItemData        =   "FormSound.frx":0876
         Left            =   5280
         List            =   "FormSound.frx":088F
         Style           =   2  '드롭다운 목록
         TabIndex        =   342
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   2820
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorGuestRegCar 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         ItemData        =   "FormSound.frx":08AF
         Left            =   5790
         List            =   "FormSound.frx":08C8
         Style           =   2  '드롭다운 목록
         TabIndex        =   341
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   2820
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp1EmgColorReg 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         ItemData        =   "FormSound.frx":08E8
         Left            =   5280
         List            =   "FormSound.frx":0901
         Style           =   2  '드롭다운 목록
         TabIndex        =   268
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   600
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorReg 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         ItemData        =   "FormSound.frx":0921
         Left            =   5790
         List            =   "FormSound.frx":093A
         Style           =   2  '드롭다운 목록
         TabIndex        =   267
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   600
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp1EmgColorGuest 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         ItemData        =   "FormSound.frx":095A
         Left            =   5280
         List            =   "FormSound.frx":0973
         Style           =   2  '드롭다운 목록
         TabIndex        =   266
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   915
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorGuest 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         ItemData        =   "FormSound.frx":0993
         Left            =   5790
         List            =   "FormSound.frx":09AC
         Style           =   2  '드롭다운 목록
         TabIndex        =   265
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   915
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp1EmgColorNoRec 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         ItemData        =   "FormSound.frx":09CC
         Left            =   5280
         List            =   "FormSound.frx":09E5
         Style           =   2  '드롭다운 목록
         TabIndex        =   264
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   1230
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorNoRec 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         ItemData        =   "FormSound.frx":0A05
         Left            =   5790
         List            =   "FormSound.frx":0A1E
         Style           =   2  '드롭다운 목록
         TabIndex        =   263
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   1230
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp1EmgColorBKList 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         ItemData        =   "FormSound.frx":0A3E
         Left            =   5280
         List            =   "FormSound.frx":0A57
         Style           =   2  '드롭다운 목록
         TabIndex        =   262
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   1545
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorBKList 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         ItemData        =   "FormSound.frx":0A77
         Left            =   5790
         List            =   "FormSound.frx":0A90
         Style           =   2  '드롭다운 목록
         TabIndex        =   261
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   1545
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp1EmgColorTaxi 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         ItemData        =   "FormSound.frx":0AB0
         Left            =   5280
         List            =   "FormSound.frx":0AC9
         Style           =   2  '드롭다운 목록
         TabIndex        =   260
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   1860
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorTaxi 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         ItemData        =   "FormSound.frx":0AE9
         Left            =   5790
         List            =   "FormSound.frx":0B02
         Style           =   2  '드롭다운 목록
         TabIndex        =   259
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   1860
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp1EmgColorDay 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         ItemData        =   "FormSound.frx":0B22
         Left            =   5280
         List            =   "FormSound.frx":0B3B
         Style           =   2  '드롭다운 목록
         TabIndex        =   258
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   2190
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorDay 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         ItemData        =   "FormSound.frx":0B5B
         Left            =   5790
         List            =   "FormSound.frx":0B74
         Style           =   2  '드롭다운 목록
         TabIndex        =   257
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   2190
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp1EmgColorRegExpDate 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         ItemData        =   "FormSound.frx":0B94
         Left            =   5280
         List            =   "FormSound.frx":0BAD
         Style           =   2  '드롭다운 목록
         TabIndex        =   256
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   2505
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorRegExpDate 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   2
         ItemData        =   "FormSound.frx":0BCD
         Left            =   5790
         List            =   "FormSound.frx":0BE6
         Style           =   2  '드롭다운 목록
         TabIndex        =   255
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   2505
         Width           =   525
      End
      Begin VB.CheckBox chk_SND_RegExpDate 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   2
         Left            =   2865
         TabIndex        =   211
         Top             =   2550
         Width           =   1200
      End
      Begin VB.CommandButton btn_PLY_RegExpDate 
         Caption         =   "▶"
         Height          =   300
         Index           =   2
         Left            =   2550
         TabIndex        =   210
         Top             =   2505
         Width           =   255
      End
      Begin VB.TextBox txt_Str_RegExpDate 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   4245
         TabIndex        =   209
         Text            =   "Text1"
         Top             =   2505
         Width           =   1020
      End
      Begin VB.CommandButton btn_SND_RegExpDate 
         Caption         =   "..."
         Height          =   300
         Index           =   2
         Left            =   2235
         TabIndex        =   208
         Top             =   2505
         Width           =   300
      End
      Begin VB.CommandButton btn_SND_Day 
         Caption         =   "..."
         Height          =   300
         Index           =   2
         Left            =   2235
         TabIndex        =   125
         Top             =   2175
         Width           =   300
      End
      Begin VB.CheckBox chk_SND_Day 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   2
         Left            =   2865
         TabIndex        =   124
         Top             =   2220
         Width           =   1200
      End
      Begin VB.TextBox txt_Str_Day 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   4245
         TabIndex        =   123
         Text            =   "Text1"
         Top             =   2175
         Width           =   1020
      End
      Begin VB.CommandButton btn_PLY_Day 
         Caption         =   "▶"
         Height          =   300
         Index           =   2
         Left            =   2550
         TabIndex        =   122
         Top             =   2175
         Width           =   255
      End
      Begin VB.CommandButton btn_PLY_Taxi 
         Caption         =   "▶"
         Height          =   300
         Index           =   2
         Left            =   2550
         TabIndex        =   121
         Top             =   1845
         Width           =   255
      End
      Begin VB.TextBox txt_Str_Taxi 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   4245
         TabIndex        =   120
         Text            =   "Text1"
         Top             =   1845
         Width           =   1020
      End
      Begin VB.CheckBox chk_SND_Taxi 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   2
         Left            =   2865
         TabIndex        =   119
         Top             =   1890
         Width           =   1200
      End
      Begin VB.CommandButton btn_SND_Taxi 
         Caption         =   "..."
         Height          =   300
         Index           =   2
         Left            =   2235
         TabIndex        =   118
         Top             =   1845
         Width           =   300
      End
      Begin VB.CheckBox chk_SND_NoRec 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   2
         Left            =   2865
         TabIndex        =   117
         Top             =   1260
         Width           =   1200
      End
      Begin VB.CheckBox chk_SND_Guest 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   2
         Left            =   2865
         TabIndex        =   116
         Top             =   945
         Width           =   1200
      End
      Begin VB.CheckBox chk_SND_BKList 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   2
         Left            =   2865
         TabIndex        =   115
         Top             =   1575
         Width           =   1200
      End
      Begin VB.TextBox txt_Str_Guest 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   4245
         TabIndex        =   114
         Text            =   "Text1"
         Top             =   900
         Width           =   1020
      End
      Begin VB.TextBox txt_Str_NoRec 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   4245
         TabIndex        =   113
         Text            =   "Text1"
         Top             =   1215
         Width           =   1020
      End
      Begin VB.TextBox txt_Str_BKList 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   4245
         TabIndex        =   112
         Text            =   "Text1"
         Top             =   1530
         Width           =   1020
      End
      Begin VB.TextBox txt_Str_Reg 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   4245
         TabIndex        =   111
         Text            =   "Text1"
         Top             =   585
         Width           =   1020
      End
      Begin VB.CommandButton btn_SND_NoRec 
         Caption         =   "..."
         Height          =   300
         Index           =   2
         Left            =   2235
         TabIndex        =   110
         Top             =   1215
         Width           =   300
      End
      Begin VB.CommandButton btn_SND_Guest 
         Caption         =   "..."
         Height          =   300
         Index           =   2
         Left            =   2235
         TabIndex        =   109
         Top             =   900
         Width           =   300
      End
      Begin VB.CommandButton btn_SND_BKList 
         Caption         =   "..."
         Height          =   300
         Index           =   2
         Left            =   2235
         TabIndex        =   108
         Top             =   1530
         Width           =   300
      End
      Begin VB.CommandButton btn_PLY_NoRec 
         Caption         =   "▶"
         Height          =   300
         Index           =   2
         Left            =   2550
         TabIndex        =   107
         Top             =   1215
         Width           =   255
      End
      Begin VB.CommandButton btn_PLY_Guest 
         Caption         =   "▶"
         Height          =   300
         Index           =   2
         Left            =   2550
         TabIndex        =   106
         Top             =   900
         Width           =   255
      End
      Begin VB.CommandButton btn_PLY_BKList 
         Caption         =   "▶"
         Height          =   300
         Index           =   2
         Left            =   2550
         TabIndex        =   105
         Top             =   1530
         Width           =   255
      End
      Begin VB.CommandButton btn_PLY_Reg 
         Caption         =   "▶"
         Height          =   300
         Index           =   2
         Left            =   2550
         TabIndex        =   104
         Top             =   585
         Width           =   255
      End
      Begin VB.CommandButton btn_SND_Reg 
         Caption         =   "..."
         Height          =   300
         Index           =   2
         Left            =   2235
         TabIndex        =   103
         Top             =   585
         Width           =   300
      End
      Begin VB.CheckBox chk_SND_Reg 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   2
         Left            =   2865
         TabIndex        =   102
         Top             =   630
         Width           =   1200
      End
      Begin VB.Label lbl_GuestRegCarExpDate 
         BackColor       =   &H00E0E0E0&
         Caption         =   "방문예약만료"
         Height          =   210
         Index           =   2
         Left            =   420
         TabIndex        =   340
         Top             =   3180
         Width           =   1335
      End
      Begin VB.Label lbl_GuestRegCar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "방문예약차량"
         Height          =   210
         Index           =   2
         Left            =   420
         TabIndex        =   339
         Top             =   2865
         Width           =   1335
      End
      Begin VB.Label lbl_RegExpDate 
         BackColor       =   &H00E0E0E0&
         Caption         =   "기간만료차량"
         Height          =   270
         Index           =   2
         Left            =   420
         TabIndex        =   207
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label21 
         BackColor       =   &H00E0E0E0&
         Caption         =   "요일제위반차량"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   420
         TabIndex        =   132
         Top             =   2190
         Width           =   1680
      End
      Begin VB.Label Label21 
         BackColor       =   &H00E0E0E0&
         Caption         =   "영업차량"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   420
         TabIndex        =   131
         Top             =   1860
         Width           =   1680
      End
      Begin VB.Label Label20 
         BackColor       =   &H00E0E0E0&
         Caption         =   "미인식차량"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   420
         TabIndex        =   130
         Top             =   1230
         Width           =   1680
      End
      Begin VB.Label Label19 
         BackColor       =   &H00E0E0E0&
         Caption         =   "미등록차량"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   420
         TabIndex        =   129
         Top             =   915
         Width           =   1680
      End
      Begin VB.Label Label18 
         BackColor       =   &H00E0E0E0&
         Caption         =   "출입제한차량"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   420
         TabIndex        =   128
         Top             =   1545
         Width           =   1680
      End
      Begin VB.Label Label19 
         BackColor       =   &H00E0E0E0&
         Caption         =   "등록차량"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   420
         TabIndex        =   127
         Top             =   600
         Width           =   1680
      End
      Begin VB.Label lbl_Lane 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00E0E0E0&
         Caption         =   "Lane3"
         BeginProperty Font 
            Name            =   "나눔고딕 ExtraBold"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   45
         TabIndex        =   126
         Top             =   165
         Width           =   675
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00E0E0E0&
      Height          =   3525
      Index           =   3
      Left            =   -30
      TabIndex        =   69
      Top             =   4635
      Width           =   6360
      Begin VB.CheckBox chk_SND_GuestRegCarExpDate 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   3
         Left            =   2865
         TabIndex        =   366
         Top             =   3180
         Width           =   1200
      End
      Begin VB.CommandButton btn_SND_GuestRegCarExpDate 
         Caption         =   "..."
         Height          =   300
         Index           =   3
         Left            =   2235
         TabIndex        =   365
         Top             =   3135
         Width           =   300
      End
      Begin VB.TextBox txt_Str_GuestRegCarExpDate 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   4245
         TabIndex        =   364
         Text            =   "Text1"
         Top             =   3135
         Width           =   1020
      End
      Begin VB.CommandButton btn_PLY_GuestRegCarExpDate 
         Caption         =   "▶"
         Height          =   300
         Index           =   3
         Left            =   2550
         TabIndex        =   363
         Top             =   3135
         Width           =   255
      End
      Begin VB.ComboBox cmb_Disp1EmgColorGuestRegCarExpDate 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   3
         ItemData        =   "FormSound.frx":0C06
         Left            =   5280
         List            =   "FormSound.frx":0C1F
         Style           =   2  '드롭다운 목록
         TabIndex        =   362
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   3135
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorGuestRegCarExpDate 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   3
         ItemData        =   "FormSound.frx":0C3F
         Left            =   5790
         List            =   "FormSound.frx":0C58
         Style           =   2  '드롭다운 목록
         TabIndex        =   361
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   3135
         Width           =   525
      End
      Begin VB.CheckBox chk_SND_GuestRegCar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   3
         Left            =   2865
         TabIndex        =   360
         Top             =   2865
         Width           =   1200
      End
      Begin VB.CommandButton btn_SND_GuestRegCar 
         Caption         =   "..."
         Height          =   300
         Index           =   3
         Left            =   2235
         TabIndex        =   359
         Top             =   2820
         Width           =   300
      End
      Begin VB.TextBox txt_Str_GuestRegCar 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   4245
         TabIndex        =   358
         Text            =   "Text1"
         Top             =   2820
         Width           =   1020
      End
      Begin VB.CommandButton btn_PLY_GuestRegCar 
         Caption         =   "▶"
         Height          =   300
         Index           =   3
         Left            =   2550
         TabIndex        =   357
         Top             =   2820
         Width           =   255
      End
      Begin VB.ComboBox cmb_Disp1EmgColorGuestRegCar 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   3
         ItemData        =   "FormSound.frx":0C78
         Left            =   5280
         List            =   "FormSound.frx":0C91
         Style           =   2  '드롭다운 목록
         TabIndex        =   356
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   2820
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorGuestRegCar 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   3
         ItemData        =   "FormSound.frx":0CB1
         Left            =   5790
         List            =   "FormSound.frx":0CCA
         Style           =   2  '드롭다운 목록
         TabIndex        =   355
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   2820
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorRegExpDate 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   3
         ItemData        =   "FormSound.frx":0CEA
         Left            =   5790
         List            =   "FormSound.frx":0D03
         Style           =   2  '드롭다운 목록
         TabIndex        =   282
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   2505
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp1EmgColorRegExpDate 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   3
         ItemData        =   "FormSound.frx":0D23
         Left            =   5280
         List            =   "FormSound.frx":0D3C
         Style           =   2  '드롭다운 목록
         TabIndex        =   281
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   2505
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorDay 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   3
         ItemData        =   "FormSound.frx":0D5C
         Left            =   5790
         List            =   "FormSound.frx":0D75
         Style           =   2  '드롭다운 목록
         TabIndex        =   280
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   2190
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp1EmgColorDay 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   3
         ItemData        =   "FormSound.frx":0D95
         Left            =   5280
         List            =   "FormSound.frx":0DAE
         Style           =   2  '드롭다운 목록
         TabIndex        =   279
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   2190
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorTaxi 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   3
         ItemData        =   "FormSound.frx":0DCE
         Left            =   5790
         List            =   "FormSound.frx":0DE7
         Style           =   2  '드롭다운 목록
         TabIndex        =   278
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   1860
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp1EmgColorTaxi 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   3
         ItemData        =   "FormSound.frx":0E07
         Left            =   5280
         List            =   "FormSound.frx":0E20
         Style           =   2  '드롭다운 목록
         TabIndex        =   277
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   1860
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorBKList 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   3
         ItemData        =   "FormSound.frx":0E40
         Left            =   5790
         List            =   "FormSound.frx":0E59
         Style           =   2  '드롭다운 목록
         TabIndex        =   276
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   1545
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp1EmgColorBKList 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   3
         ItemData        =   "FormSound.frx":0E79
         Left            =   5280
         List            =   "FormSound.frx":0E92
         Style           =   2  '드롭다운 목록
         TabIndex        =   275
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   1545
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorNoRec 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   3
         ItemData        =   "FormSound.frx":0EB2
         Left            =   5790
         List            =   "FormSound.frx":0ECB
         Style           =   2  '드롭다운 목록
         TabIndex        =   274
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   1230
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp1EmgColorNoRec 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   3
         ItemData        =   "FormSound.frx":0EEB
         Left            =   5280
         List            =   "FormSound.frx":0F04
         Style           =   2  '드롭다운 목록
         TabIndex        =   273
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   1230
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorGuest 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   3
         ItemData        =   "FormSound.frx":0F24
         Left            =   5790
         List            =   "FormSound.frx":0F3D
         Style           =   2  '드롭다운 목록
         TabIndex        =   272
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   915
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp1EmgColorGuest 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   3
         ItemData        =   "FormSound.frx":0F5D
         Left            =   5280
         List            =   "FormSound.frx":0F76
         Style           =   2  '드롭다운 목록
         TabIndex        =   271
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   915
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorReg 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   3
         ItemData        =   "FormSound.frx":0F96
         Left            =   5790
         List            =   "FormSound.frx":0FAF
         Style           =   2  '드롭다운 목록
         TabIndex        =   270
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   600
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp1EmgColorReg 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   3
         ItemData        =   "FormSound.frx":0FCF
         Left            =   5280
         List            =   "FormSound.frx":0FE8
         Style           =   2  '드롭다운 목록
         TabIndex        =   269
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   600
         Width           =   525
      End
      Begin VB.CheckBox chk_SND_RegExpDate 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   3
         Left            =   2865
         TabIndex        =   216
         Top             =   2550
         Width           =   1200
      End
      Begin VB.CommandButton btn_PLY_RegExpDate 
         Caption         =   "▶"
         Height          =   300
         Index           =   3
         Left            =   2550
         TabIndex        =   215
         Top             =   2505
         Width           =   255
      End
      Begin VB.TextBox txt_Str_RegExpDate 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   4245
         TabIndex        =   214
         Text            =   "Text1"
         Top             =   2505
         Width           =   1020
      End
      Begin VB.CommandButton btn_SND_RegExpDate 
         Caption         =   "..."
         Height          =   300
         Index           =   3
         Left            =   2235
         TabIndex        =   213
         Top             =   2505
         Width           =   300
      End
      Begin VB.CheckBox chk_SND_Reg 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   3
         Left            =   2865
         TabIndex        =   93
         Top             =   630
         Width           =   1200
      End
      Begin VB.CommandButton btn_SND_Reg 
         Caption         =   "..."
         Height          =   300
         Index           =   3
         Left            =   2235
         TabIndex        =   92
         Top             =   585
         Width           =   300
      End
      Begin VB.CommandButton btn_PLY_Reg 
         Caption         =   "▶"
         Height          =   300
         Index           =   3
         Left            =   2550
         TabIndex        =   91
         Top             =   585
         Width           =   255
      End
      Begin VB.CommandButton btn_PLY_BKList 
         Caption         =   "▶"
         Height          =   300
         Index           =   3
         Left            =   2550
         TabIndex        =   90
         Top             =   1530
         Width           =   255
      End
      Begin VB.CommandButton btn_PLY_Guest 
         Caption         =   "▶"
         Height          =   300
         Index           =   3
         Left            =   2550
         TabIndex        =   89
         Top             =   900
         Width           =   255
      End
      Begin VB.CommandButton btn_PLY_NoRec 
         Caption         =   "▶"
         Height          =   300
         Index           =   3
         Left            =   2550
         TabIndex        =   88
         Top             =   1215
         Width           =   255
      End
      Begin VB.CommandButton btn_SND_BKList 
         Caption         =   "..."
         Height          =   300
         Index           =   3
         Left            =   2235
         TabIndex        =   87
         Top             =   1530
         Width           =   300
      End
      Begin VB.CommandButton btn_SND_Guest 
         Caption         =   "..."
         Height          =   300
         Index           =   3
         Left            =   2235
         TabIndex        =   86
         Top             =   900
         Width           =   300
      End
      Begin VB.CommandButton btn_SND_NoRec 
         Caption         =   "..."
         Height          =   300
         Index           =   3
         Left            =   2235
         TabIndex        =   85
         Top             =   1215
         Width           =   300
      End
      Begin VB.TextBox txt_Str_Reg 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   4245
         TabIndex        =   84
         Text            =   "Text1"
         Top             =   585
         Width           =   1020
      End
      Begin VB.TextBox txt_Str_BKList 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   4245
         TabIndex        =   83
         Text            =   "Text1"
         Top             =   1530
         Width           =   1020
      End
      Begin VB.TextBox txt_Str_NoRec 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   4245
         TabIndex        =   82
         Text            =   "Text1"
         Top             =   1215
         Width           =   1020
      End
      Begin VB.TextBox txt_Str_Guest 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   4245
         TabIndex        =   81
         Text            =   "Text1"
         Top             =   900
         Width           =   1020
      End
      Begin VB.CheckBox chk_SND_BKList 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   3
         Left            =   2865
         TabIndex        =   80
         Top             =   1575
         Width           =   1200
      End
      Begin VB.CheckBox chk_SND_Guest 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   3
         Left            =   2865
         TabIndex        =   79
         Top             =   945
         Width           =   1200
      End
      Begin VB.CheckBox chk_SND_NoRec 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   3
         Left            =   2865
         TabIndex        =   78
         Top             =   1260
         Width           =   1200
      End
      Begin VB.CommandButton btn_SND_Taxi 
         Caption         =   "..."
         Height          =   300
         Index           =   3
         Left            =   2235
         TabIndex        =   77
         Top             =   1845
         Width           =   300
      End
      Begin VB.CheckBox chk_SND_Taxi 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   3
         Left            =   2865
         TabIndex        =   76
         Top             =   1890
         Width           =   1200
      End
      Begin VB.TextBox txt_Str_Taxi 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   4245
         TabIndex        =   75
         Text            =   "Text1"
         Top             =   1845
         Width           =   1020
      End
      Begin VB.CommandButton btn_PLY_Taxi 
         Caption         =   "▶"
         Height          =   300
         Index           =   3
         Left            =   2550
         TabIndex        =   74
         Top             =   1845
         Width           =   255
      End
      Begin VB.CommandButton btn_PLY_Day 
         Caption         =   "▶"
         Height          =   300
         Index           =   3
         Left            =   2550
         TabIndex        =   73
         Top             =   2175
         Width           =   255
      End
      Begin VB.TextBox txt_Str_Day 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   4245
         TabIndex        =   72
         Text            =   "Text1"
         Top             =   2175
         Width           =   1020
      End
      Begin VB.CheckBox chk_SND_Day 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   3
         Left            =   2865
         TabIndex        =   71
         Top             =   2220
         Width           =   1200
      End
      Begin VB.CommandButton btn_SND_Day 
         Caption         =   "..."
         Height          =   300
         Index           =   3
         Left            =   2235
         TabIndex        =   70
         Top             =   2175
         Width           =   300
      End
      Begin VB.Label lbl_GuestRegCarExpDate 
         BackColor       =   &H00E0E0E0&
         Caption         =   "방문예약만료"
         Height          =   210
         Index           =   3
         Left            =   420
         TabIndex        =   354
         Top             =   3180
         Width           =   1335
      End
      Begin VB.Label lbl_GuestRegCar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "방문예약차량"
         Height          =   210
         Index           =   3
         Left            =   420
         TabIndex        =   353
         Top             =   2865
         Width           =   1335
      End
      Begin VB.Label lbl_RegExpDate 
         BackColor       =   &H00E0E0E0&
         Caption         =   "기간만료차량"
         Height          =   270
         Index           =   3
         Left            =   420
         TabIndex        =   212
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label lbl_Lane 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00E0E0E0&
         Caption         =   "Lane4"
         BeginProperty Font 
            Name            =   "나눔고딕 ExtraBold"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   45
         TabIndex        =   100
         Top             =   165
         Width           =   675
      End
      Begin VB.Label Label19 
         BackColor       =   &H00E0E0E0&
         Caption         =   "등록차량"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   420
         TabIndex        =   99
         Top             =   600
         Width           =   1680
      End
      Begin VB.Label Label18 
         BackColor       =   &H00E0E0E0&
         Caption         =   "출입제한차량"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   420
         TabIndex        =   98
         Top             =   1545
         Width           =   1680
      End
      Begin VB.Label Label19 
         BackColor       =   &H00E0E0E0&
         Caption         =   "미등록차량"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   420
         TabIndex        =   97
         Top             =   915
         Width           =   1680
      End
      Begin VB.Label Label20 
         BackColor       =   &H00E0E0E0&
         Caption         =   "미인식차량"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   420
         TabIndex        =   96
         Top             =   1230
         Width           =   1680
      End
      Begin VB.Label Label21 
         BackColor       =   &H00E0E0E0&
         Caption         =   "영업차량"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   420
         TabIndex        =   95
         Top             =   1860
         Width           =   1680
      End
      Begin VB.Label Label21 
         BackColor       =   &H00E0E0E0&
         Caption         =   "요일제위반차량"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   420
         TabIndex        =   94
         Top             =   2190
         Width           =   1680
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00E0E0E0&
      Height          =   3525
      Index           =   4
      Left            =   6465
      TabIndex        =   37
      Top             =   4635
      Width           =   6360
      Begin VB.CheckBox chk_SND_GuestRegCarExpDate 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   4
         Left            =   2865
         TabIndex        =   380
         Top             =   3180
         Width           =   1200
      End
      Begin VB.CommandButton btn_SND_GuestRegCarExpDate 
         Caption         =   "..."
         Height          =   300
         Index           =   4
         Left            =   2235
         TabIndex        =   379
         Top             =   3135
         Width           =   300
      End
      Begin VB.TextBox txt_Str_GuestRegCarExpDate 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   4245
         TabIndex        =   378
         Text            =   "Text1"
         Top             =   3135
         Width           =   1020
      End
      Begin VB.CommandButton btn_PLY_GuestRegCarExpDate 
         Caption         =   "▶"
         Height          =   300
         Index           =   4
         Left            =   2550
         TabIndex        =   377
         Top             =   3135
         Width           =   255
      End
      Begin VB.ComboBox cmb_Disp1EmgColorGuestRegCarExpDate 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   4
         ItemData        =   "FormSound.frx":1008
         Left            =   5280
         List            =   "FormSound.frx":1021
         Style           =   2  '드롭다운 목록
         TabIndex        =   376
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   3135
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorGuestRegCarExpDate 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   4
         ItemData        =   "FormSound.frx":1041
         Left            =   5790
         List            =   "FormSound.frx":105A
         Style           =   2  '드롭다운 목록
         TabIndex        =   375
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   3135
         Width           =   525
      End
      Begin VB.CheckBox chk_SND_GuestRegCar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   4
         Left            =   2865
         TabIndex        =   374
         Top             =   2865
         Width           =   1200
      End
      Begin VB.CommandButton btn_SND_GuestRegCar 
         Caption         =   "..."
         Height          =   300
         Index           =   4
         Left            =   2235
         TabIndex        =   373
         Top             =   2820
         Width           =   300
      End
      Begin VB.TextBox txt_Str_GuestRegCar 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   4245
         TabIndex        =   372
         Text            =   "Text1"
         Top             =   2820
         Width           =   1020
      End
      Begin VB.CommandButton btn_PLY_GuestRegCar 
         Caption         =   "▶"
         Height          =   300
         Index           =   4
         Left            =   2550
         TabIndex        =   371
         Top             =   2820
         Width           =   255
      End
      Begin VB.ComboBox cmb_Disp1EmgColorGuestRegCar 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   4
         ItemData        =   "FormSound.frx":107A
         Left            =   5280
         List            =   "FormSound.frx":1093
         Style           =   2  '드롭다운 목록
         TabIndex        =   370
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   2820
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorGuestRegCar 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   4
         ItemData        =   "FormSound.frx":10B3
         Left            =   5790
         List            =   "FormSound.frx":10CC
         Style           =   2  '드롭다운 목록
         TabIndex        =   369
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   2820
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorRegExpDate 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   4
         ItemData        =   "FormSound.frx":10EC
         Left            =   5790
         List            =   "FormSound.frx":1105
         Style           =   2  '드롭다운 목록
         TabIndex        =   296
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   2505
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp1EmgColorRegExpDate 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   4
         ItemData        =   "FormSound.frx":1125
         Left            =   5280
         List            =   "FormSound.frx":113E
         Style           =   2  '드롭다운 목록
         TabIndex        =   295
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   2505
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorDay 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   4
         ItemData        =   "FormSound.frx":115E
         Left            =   5790
         List            =   "FormSound.frx":1177
         Style           =   2  '드롭다운 목록
         TabIndex        =   294
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   2190
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp1EmgColorDay 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   4
         ItemData        =   "FormSound.frx":1197
         Left            =   5280
         List            =   "FormSound.frx":11B0
         Style           =   2  '드롭다운 목록
         TabIndex        =   293
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   2190
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorTaxi 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   4
         ItemData        =   "FormSound.frx":11D0
         Left            =   5790
         List            =   "FormSound.frx":11E9
         Style           =   2  '드롭다운 목록
         TabIndex        =   292
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   1860
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp1EmgColorTaxi 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   4
         ItemData        =   "FormSound.frx":1209
         Left            =   5280
         List            =   "FormSound.frx":1222
         Style           =   2  '드롭다운 목록
         TabIndex        =   291
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   1860
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorBKList 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   4
         ItemData        =   "FormSound.frx":1242
         Left            =   5790
         List            =   "FormSound.frx":125B
         Style           =   2  '드롭다운 목록
         TabIndex        =   290
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   1545
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp1EmgColorBKList 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   4
         ItemData        =   "FormSound.frx":127B
         Left            =   5280
         List            =   "FormSound.frx":1294
         Style           =   2  '드롭다운 목록
         TabIndex        =   289
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   1545
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorNoRec 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   4
         ItemData        =   "FormSound.frx":12B4
         Left            =   5790
         List            =   "FormSound.frx":12CD
         Style           =   2  '드롭다운 목록
         TabIndex        =   288
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   1230
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp1EmgColorNoRec 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   4
         ItemData        =   "FormSound.frx":12ED
         Left            =   5280
         List            =   "FormSound.frx":1306
         Style           =   2  '드롭다운 목록
         TabIndex        =   287
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   1230
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorGuest 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   4
         ItemData        =   "FormSound.frx":1326
         Left            =   5790
         List            =   "FormSound.frx":133F
         Style           =   2  '드롭다운 목록
         TabIndex        =   286
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   915
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp1EmgColorGuest 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   4
         ItemData        =   "FormSound.frx":135F
         Left            =   5280
         List            =   "FormSound.frx":1378
         Style           =   2  '드롭다운 목록
         TabIndex        =   285
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   915
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorReg 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   4
         ItemData        =   "FormSound.frx":1398
         Left            =   5790
         List            =   "FormSound.frx":13B1
         Style           =   2  '드롭다운 목록
         TabIndex        =   284
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   600
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp1EmgColorReg 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   4
         ItemData        =   "FormSound.frx":13D1
         Left            =   5280
         List            =   "FormSound.frx":13EA
         Style           =   2  '드롭다운 목록
         TabIndex        =   283
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   600
         Width           =   525
      End
      Begin VB.CheckBox chk_SND_RegExpDate 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   4
         Left            =   2865
         TabIndex        =   221
         Top             =   2550
         Width           =   1200
      End
      Begin VB.CommandButton btn_PLY_RegExpDate 
         Caption         =   "▶"
         Height          =   300
         Index           =   4
         Left            =   2550
         TabIndex        =   220
         Top             =   2505
         Width           =   255
      End
      Begin VB.TextBox txt_Str_RegExpDate 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   4245
         TabIndex        =   219
         Text            =   "Text1"
         Top             =   2505
         Width           =   1020
      End
      Begin VB.CommandButton btn_SND_RegExpDate 
         Caption         =   "..."
         Height          =   300
         Index           =   4
         Left            =   2235
         TabIndex        =   218
         Top             =   2505
         Width           =   300
      End
      Begin VB.CommandButton btn_SND_Day 
         Caption         =   "..."
         Height          =   300
         Index           =   4
         Left            =   2235
         TabIndex        =   61
         Top             =   2175
         Width           =   300
      End
      Begin VB.CheckBox chk_SND_Day 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   4
         Left            =   2865
         TabIndex        =   60
         Top             =   2220
         Width           =   1200
      End
      Begin VB.TextBox txt_Str_Day 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   4245
         TabIndex        =   59
         Text            =   "Text1"
         Top             =   2175
         Width           =   1020
      End
      Begin VB.CommandButton btn_PLY_Day 
         Caption         =   "▶"
         Height          =   300
         Index           =   4
         Left            =   2550
         TabIndex        =   58
         Top             =   2175
         Width           =   255
      End
      Begin VB.CommandButton btn_PLY_Taxi 
         Caption         =   "▶"
         Height          =   300
         Index           =   4
         Left            =   2550
         TabIndex        =   57
         Top             =   1845
         Width           =   255
      End
      Begin VB.TextBox txt_Str_Taxi 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   4245
         TabIndex        =   56
         Text            =   "Text1"
         Top             =   1845
         Width           =   1020
      End
      Begin VB.CheckBox chk_SND_Taxi 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   4
         Left            =   2865
         TabIndex        =   55
         Top             =   1890
         Width           =   1200
      End
      Begin VB.CommandButton btn_SND_Taxi 
         Caption         =   "..."
         Height          =   300
         Index           =   4
         Left            =   2235
         TabIndex        =   54
         Top             =   1845
         Width           =   300
      End
      Begin VB.CheckBox chk_SND_NoRec 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   4
         Left            =   2865
         TabIndex        =   53
         Top             =   1260
         Width           =   1200
      End
      Begin VB.CheckBox chk_SND_Guest 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   4
         Left            =   2865
         TabIndex        =   52
         Top             =   945
         Width           =   1200
      End
      Begin VB.CheckBox chk_SND_BKList 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   4
         Left            =   2865
         TabIndex        =   51
         Top             =   1575
         Width           =   1200
      End
      Begin VB.TextBox txt_Str_Guest 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   4245
         TabIndex        =   50
         Text            =   "Text1"
         Top             =   900
         Width           =   1020
      End
      Begin VB.TextBox txt_Str_NoRec 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   4245
         TabIndex        =   49
         Text            =   "Text1"
         Top             =   1215
         Width           =   1020
      End
      Begin VB.TextBox txt_Str_BKList 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   4245
         TabIndex        =   48
         Text            =   "Text1"
         Top             =   1530
         Width           =   1020
      End
      Begin VB.TextBox txt_Str_Reg 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   4245
         TabIndex        =   47
         Text            =   "Text1"
         Top             =   585
         Width           =   1020
      End
      Begin VB.CommandButton btn_SND_NoRec 
         Caption         =   "..."
         Height          =   300
         Index           =   4
         Left            =   2235
         TabIndex        =   46
         Top             =   1215
         Width           =   300
      End
      Begin VB.CommandButton btn_SND_Guest 
         Caption         =   "..."
         Height          =   300
         Index           =   4
         Left            =   2235
         TabIndex        =   45
         Top             =   900
         Width           =   300
      End
      Begin VB.CommandButton btn_SND_BKList 
         Caption         =   "..."
         Height          =   300
         Index           =   4
         Left            =   2235
         TabIndex        =   44
         Top             =   1530
         Width           =   300
      End
      Begin VB.CommandButton btn_PLY_NoRec 
         Caption         =   "▶"
         Height          =   300
         Index           =   4
         Left            =   2550
         TabIndex        =   43
         Top             =   1215
         Width           =   255
      End
      Begin VB.CommandButton btn_PLY_Guest 
         Caption         =   "▶"
         Height          =   300
         Index           =   4
         Left            =   2550
         TabIndex        =   42
         Top             =   900
         Width           =   255
      End
      Begin VB.CommandButton btn_PLY_BKList 
         Caption         =   "▶"
         Height          =   300
         Index           =   4
         Left            =   2550
         TabIndex        =   41
         Top             =   1530
         Width           =   255
      End
      Begin VB.CommandButton btn_PLY_Reg 
         Caption         =   "▶"
         Height          =   300
         Index           =   4
         Left            =   2550
         TabIndex        =   40
         Top             =   585
         Width           =   255
      End
      Begin VB.CommandButton btn_SND_Reg 
         Caption         =   "..."
         Height          =   300
         Index           =   4
         Left            =   2235
         TabIndex        =   39
         Top             =   585
         Width           =   300
      End
      Begin VB.CheckBox chk_SND_Reg 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   4
         Left            =   2865
         TabIndex        =   38
         Top             =   630
         Width           =   1200
      End
      Begin VB.Label lbl_GuestRegCarExpDate 
         BackColor       =   &H00E0E0E0&
         Caption         =   "방문예약만료"
         Height          =   210
         Index           =   4
         Left            =   420
         TabIndex        =   368
         Top             =   3180
         Width           =   1335
      End
      Begin VB.Label lbl_GuestRegCar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "방문예약차량"
         Height          =   210
         Index           =   4
         Left            =   420
         TabIndex        =   367
         Top             =   2865
         Width           =   1335
      End
      Begin VB.Label lbl_RegExpDate 
         BackColor       =   &H00E0E0E0&
         Caption         =   "기간만료차량"
         Height          =   270
         Index           =   4
         Left            =   420
         TabIndex        =   217
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label21 
         BackColor       =   &H00E0E0E0&
         Caption         =   "요일제위반차량"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   9
         Left            =   420
         TabIndex        =   68
         Top             =   2190
         Width           =   1680
      End
      Begin VB.Label Label21 
         BackColor       =   &H00E0E0E0&
         Caption         =   "영업차량"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   420
         TabIndex        =   67
         Top             =   1860
         Width           =   1680
      End
      Begin VB.Label Label20 
         BackColor       =   &H00E0E0E0&
         Caption         =   "미인식차량"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   420
         TabIndex        =   66
         Top             =   1230
         Width           =   1680
      End
      Begin VB.Label Label19 
         BackColor       =   &H00E0E0E0&
         Caption         =   "미등록차량"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   9
         Left            =   420
         TabIndex        =   65
         Top             =   915
         Width           =   1680
      End
      Begin VB.Label Label18 
         BackColor       =   &H00E0E0E0&
         Caption         =   "출입제한차량"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   420
         TabIndex        =   64
         Top             =   1545
         Width           =   1680
      End
      Begin VB.Label Label19 
         BackColor       =   &H00E0E0E0&
         Caption         =   "등록차량"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   420
         TabIndex        =   63
         Top             =   600
         Width           =   1680
      End
      Begin VB.Label lbl_Lane 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00E0E0E0&
         Caption         =   "Lane5"
         BeginProperty Font 
            Name            =   "나눔고딕 ExtraBold"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   45
         TabIndex        =   62
         Top             =   165
         Width           =   675
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00E0E0E0&
      Height          =   3525
      Index           =   5
      Left            =   12960
      TabIndex        =   5
      Top             =   4635
      Width           =   6360
      Begin VB.CheckBox chk_SND_GuestRegCarExpDate 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   5
         Left            =   2865
         TabIndex        =   394
         Top             =   3165
         Width           =   1200
      End
      Begin VB.CommandButton btn_SND_GuestRegCarExpDate 
         Caption         =   "..."
         Height          =   300
         Index           =   5
         Left            =   2235
         TabIndex        =   393
         Top             =   3120
         Width           =   300
      End
      Begin VB.TextBox txt_Str_GuestRegCarExpDate 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   4245
         TabIndex        =   392
         Text            =   "Text1"
         Top             =   3120
         Width           =   1020
      End
      Begin VB.CommandButton btn_PLY_GuestRegCarExpDate 
         Caption         =   "▶"
         Height          =   300
         Index           =   5
         Left            =   2550
         TabIndex        =   391
         Top             =   3120
         Width           =   255
      End
      Begin VB.ComboBox cmb_Disp1EmgColorGuestRegCarExpDate 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   5
         ItemData        =   "FormSound.frx":140A
         Left            =   5280
         List            =   "FormSound.frx":1423
         Style           =   2  '드롭다운 목록
         TabIndex        =   390
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   3120
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorGuestRegCarExpDate 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   5
         ItemData        =   "FormSound.frx":1443
         Left            =   5790
         List            =   "FormSound.frx":145C
         Style           =   2  '드롭다운 목록
         TabIndex        =   389
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   3120
         Width           =   525
      End
      Begin VB.CheckBox chk_SND_GuestRegCar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   5
         Left            =   2865
         TabIndex        =   388
         Top             =   2850
         Width           =   1200
      End
      Begin VB.CommandButton btn_SND_GuestRegCar 
         Caption         =   "..."
         Height          =   300
         Index           =   5
         Left            =   2235
         TabIndex        =   387
         Top             =   2805
         Width           =   300
      End
      Begin VB.TextBox txt_Str_GuestRegCar 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   4245
         TabIndex        =   386
         Text            =   "Text1"
         Top             =   2805
         Width           =   1020
      End
      Begin VB.CommandButton btn_PLY_GuestRegCar 
         Caption         =   "▶"
         Height          =   300
         Index           =   5
         Left            =   2550
         TabIndex        =   385
         Top             =   2805
         Width           =   255
      End
      Begin VB.ComboBox cmb_Disp1EmgColorGuestRegCar 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   5
         ItemData        =   "FormSound.frx":147C
         Left            =   5280
         List            =   "FormSound.frx":1495
         Style           =   2  '드롭다운 목록
         TabIndex        =   384
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   2805
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorGuestRegCar 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   5
         ItemData        =   "FormSound.frx":14B5
         Left            =   5790
         List            =   "FormSound.frx":14CE
         Style           =   2  '드롭다운 목록
         TabIndex        =   383
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   2805
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorRegExpDate 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   5
         ItemData        =   "FormSound.frx":14EE
         Left            =   5790
         List            =   "FormSound.frx":1507
         Style           =   2  '드롭다운 목록
         TabIndex        =   310
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   2490
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp1EmgColorRegExpDate 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   5
         ItemData        =   "FormSound.frx":1527
         Left            =   5280
         List            =   "FormSound.frx":1540
         Style           =   2  '드롭다운 목록
         TabIndex        =   309
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   2490
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorDay 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   5
         ItemData        =   "FormSound.frx":1560
         Left            =   5790
         List            =   "FormSound.frx":1579
         Style           =   2  '드롭다운 목록
         TabIndex        =   308
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   2175
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp1EmgColorDay 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   5
         ItemData        =   "FormSound.frx":1599
         Left            =   5280
         List            =   "FormSound.frx":15B2
         Style           =   2  '드롭다운 목록
         TabIndex        =   307
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   2175
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorTaxi 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   5
         ItemData        =   "FormSound.frx":15D2
         Left            =   5790
         List            =   "FormSound.frx":15EB
         Style           =   2  '드롭다운 목록
         TabIndex        =   306
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   1845
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp1EmgColorTaxi 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   5
         ItemData        =   "FormSound.frx":160B
         Left            =   5280
         List            =   "FormSound.frx":1624
         Style           =   2  '드롭다운 목록
         TabIndex        =   305
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   1845
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorBKList 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   5
         ItemData        =   "FormSound.frx":1644
         Left            =   5790
         List            =   "FormSound.frx":165D
         Style           =   2  '드롭다운 목록
         TabIndex        =   304
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   1530
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp1EmgColorBKList 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   5
         ItemData        =   "FormSound.frx":167D
         Left            =   5280
         List            =   "FormSound.frx":1696
         Style           =   2  '드롭다운 목록
         TabIndex        =   303
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   1530
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorNoRec 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   5
         ItemData        =   "FormSound.frx":16B6
         Left            =   5790
         List            =   "FormSound.frx":16CF
         Style           =   2  '드롭다운 목록
         TabIndex        =   302
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   1215
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp1EmgColorNoRec 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   5
         ItemData        =   "FormSound.frx":16EF
         Left            =   5280
         List            =   "FormSound.frx":1708
         Style           =   2  '드롭다운 목록
         TabIndex        =   301
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   1215
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorGuest 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   5
         ItemData        =   "FormSound.frx":1728
         Left            =   5790
         List            =   "FormSound.frx":1741
         Style           =   2  '드롭다운 목록
         TabIndex        =   300
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   900
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp1EmgColorGuest 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   5
         ItemData        =   "FormSound.frx":1761
         Left            =   5280
         List            =   "FormSound.frx":177A
         Style           =   2  '드롭다운 목록
         TabIndex        =   299
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   900
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp2EmgColorReg 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   5
         ItemData        =   "FormSound.frx":179A
         Left            =   5790
         List            =   "FormSound.frx":17B3
         Style           =   2  '드롭다운 목록
         TabIndex        =   298
         ToolTipText     =   "전광판 둘째줄 색상"
         Top             =   585
         Width           =   525
      End
      Begin VB.ComboBox cmb_Disp1EmgColorReg 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   5
         ItemData        =   "FormSound.frx":17D3
         Left            =   5280
         List            =   "FormSound.frx":17EC
         Style           =   2  '드롭다운 목록
         TabIndex        =   297
         ToolTipText     =   "전광판 첫째줄 색상"
         Top             =   585
         Width           =   525
      End
      Begin VB.CheckBox chk_SND_RegExpDate 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   5
         Left            =   2865
         TabIndex        =   226
         Top             =   2535
         Width           =   1200
      End
      Begin VB.CommandButton btn_PLY_RegExpDate 
         Caption         =   "▶"
         Height          =   300
         Index           =   5
         Left            =   2550
         TabIndex        =   225
         Top             =   2490
         Width           =   255
      End
      Begin VB.TextBox txt_Str_RegExpDate 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   4245
         TabIndex        =   224
         Text            =   "Text1"
         Top             =   2490
         Width           =   1020
      End
      Begin VB.CommandButton btn_SND_RegExpDate 
         Caption         =   "..."
         Height          =   300
         Index           =   5
         Left            =   2235
         TabIndex        =   223
         Top             =   2490
         Width           =   300
      End
      Begin VB.CheckBox chk_SND_Reg 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   5
         Left            =   2865
         TabIndex        =   29
         Top             =   630
         Width           =   1200
      End
      Begin VB.CommandButton btn_SND_Reg 
         Caption         =   "..."
         Height          =   300
         Index           =   5
         Left            =   2235
         TabIndex        =   28
         Top             =   585
         Width           =   300
      End
      Begin VB.CommandButton btn_PLY_Reg 
         Caption         =   "▶"
         Height          =   300
         Index           =   5
         Left            =   2550
         TabIndex        =   27
         Top             =   585
         Width           =   255
      End
      Begin VB.CommandButton btn_PLY_BKList 
         Caption         =   "▶"
         Height          =   300
         Index           =   5
         Left            =   2550
         TabIndex        =   26
         Top             =   1530
         Width           =   255
      End
      Begin VB.CommandButton btn_PLY_Guest 
         Caption         =   "▶"
         Height          =   300
         Index           =   5
         Left            =   2550
         TabIndex        =   25
         Top             =   900
         Width           =   255
      End
      Begin VB.CommandButton btn_PLY_NoRec 
         Caption         =   "▶"
         Height          =   300
         Index           =   5
         Left            =   2550
         TabIndex        =   24
         Top             =   1215
         Width           =   255
      End
      Begin VB.CommandButton btn_SND_BKList 
         Caption         =   "..."
         Height          =   300
         Index           =   5
         Left            =   2235
         TabIndex        =   23
         Top             =   1530
         Width           =   300
      End
      Begin VB.CommandButton btn_SND_Guest 
         Caption         =   "..."
         Height          =   300
         Index           =   5
         Left            =   2235
         TabIndex        =   22
         Top             =   900
         Width           =   300
      End
      Begin VB.CommandButton btn_SND_NoRec 
         Caption         =   "..."
         Height          =   300
         Index           =   5
         Left            =   2235
         TabIndex        =   21
         Top             =   1215
         Width           =   300
      End
      Begin VB.TextBox txt_Str_Reg 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   4245
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   585
         Width           =   1020
      End
      Begin VB.TextBox txt_Str_BKList 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   4245
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   1530
         Width           =   1020
      End
      Begin VB.TextBox txt_Str_NoRec 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   4245
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   1215
         Width           =   1020
      End
      Begin VB.TextBox txt_Str_Guest 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   4245
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   900
         Width           =   1020
      End
      Begin VB.CheckBox chk_SND_BKList 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   5
         Left            =   2865
         TabIndex        =   16
         Top             =   1575
         Width           =   1200
      End
      Begin VB.CheckBox chk_SND_Guest 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   5
         Left            =   2865
         TabIndex        =   15
         Top             =   945
         Width           =   1200
      End
      Begin VB.CheckBox chk_SND_NoRec 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   5
         Left            =   2865
         TabIndex        =   14
         Top             =   1260
         Width           =   1200
      End
      Begin VB.CommandButton btn_SND_Taxi 
         Caption         =   "..."
         Height          =   300
         Index           =   5
         Left            =   2235
         TabIndex        =   13
         Top             =   1845
         Width           =   300
      End
      Begin VB.CheckBox chk_SND_Taxi 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   5
         Left            =   2865
         TabIndex        =   12
         Top             =   1890
         Width           =   1200
      End
      Begin VB.TextBox txt_Str_Taxi 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   4245
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1845
         Width           =   1020
      End
      Begin VB.CommandButton btn_PLY_Taxi 
         Caption         =   "▶"
         Height          =   300
         Index           =   5
         Left            =   2550
         TabIndex        =   10
         Top             =   1845
         Width           =   255
      End
      Begin VB.CommandButton btn_PLY_Day 
         Caption         =   "▶"
         Height          =   300
         Index           =   5
         Left            =   2550
         TabIndex        =   9
         Top             =   2175
         Width           =   255
      End
      Begin VB.TextBox txt_Str_Day 
         BeginProperty Font 
            Name            =   "돋움체"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   4245
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   2175
         Width           =   1020
      End
      Begin VB.CheckBox chk_SND_Day 
         BackColor       =   &H00E0E0E0&
         Caption         =   "사운드사용"
         Height          =   225
         Index           =   5
         Left            =   2865
         TabIndex        =   7
         Top             =   2220
         Width           =   1200
      End
      Begin VB.CommandButton btn_SND_Day 
         Caption         =   "..."
         Height          =   300
         Index           =   5
         Left            =   2235
         TabIndex        =   6
         Top             =   2175
         Width           =   300
      End
      Begin VB.Label lbl_GuestRegCarExpDate 
         BackColor       =   &H00E0E0E0&
         Caption         =   "방문예약만료"
         Height          =   210
         Index           =   5
         Left            =   420
         TabIndex        =   382
         Top             =   3165
         Width           =   1335
      End
      Begin VB.Label lbl_GuestRegCar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "방문예약차량"
         Height          =   210
         Index           =   5
         Left            =   420
         TabIndex        =   381
         Top             =   2850
         Width           =   1335
      End
      Begin VB.Label lbl_RegExpDate 
         BackColor       =   &H00E0E0E0&
         Caption         =   "기간만료차량"
         Height          =   270
         Index           =   5
         Left            =   420
         TabIndex        =   222
         Top             =   2505
         Width           =   1335
      End
      Begin VB.Label lbl_Lane 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00E0E0E0&
         Caption         =   "Lane6"
         BeginProperty Font 
            Name            =   "나눔고딕 ExtraBold"
            Size            =   9.75
            Charset         =   129
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   45
         TabIndex        =   36
         Top             =   165
         Width           =   675
      End
      Begin VB.Label Label19 
         BackColor       =   &H00E0E0E0&
         Caption         =   "등록차량"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   11
         Left            =   420
         TabIndex        =   35
         Top             =   600
         Width           =   1680
      End
      Begin VB.Label Label18 
         BackColor       =   &H00E0E0E0&
         Caption         =   "출입제한차량"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   420
         TabIndex        =   34
         Top             =   1545
         Width           =   1680
      End
      Begin VB.Label Label19 
         BackColor       =   &H00E0E0E0&
         Caption         =   "미등록차량"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   10
         Left            =   420
         TabIndex        =   33
         Top             =   915
         Width           =   1680
      End
      Begin VB.Label Label20 
         BackColor       =   &H00E0E0E0&
         Caption         =   "미인식차량"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   420
         TabIndex        =   32
         Top             =   1230
         Width           =   1680
      End
      Begin VB.Label Label21 
         BackColor       =   &H00E0E0E0&
         Caption         =   "영업차량"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   11
         Left            =   420
         TabIndex        =   31
         Top             =   1860
         Width           =   1680
      End
      Begin VB.Label Label21 
         BackColor       =   &H00E0E0E0&
         Caption         =   "요일제위반차량"
         BeginProperty Font 
            Name            =   "굴림"
            Size            =   9.75
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   10
         Left            =   420
         TabIndex        =   30
         Top             =   2190
         Width           =   1680
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "닫기"
      BeginProperty Font 
         Name            =   "나눔고딕 ExtraBold"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   17505
      TabIndex        =   3
      Top             =   225
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "저장"
      BeginProperty Font 
         Name            =   "나눔고딕 ExtraBold"
         Size            =   12
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   15855
      TabIndex        =   2
      Top             =   225
      Width           =   1455
   End
   Begin VB.CheckBox chk_SOUND_YN 
      BackColor       =   &H00E0E0E0&
      Caption         =   "사운드사용"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   11115
      TabIndex        =   1
      Top             =   300
      Width           =   1530
   End
   Begin VB.CommandButton cmd_Snd_Default 
      Caption         =   "기본설정"
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   12975
      TabIndex        =   0
      ToolTipText     =   "사운드사용 및 전광판문구 기본설정"
      Top             =   225
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   " 사운드 및 전광판 긴급문구 설정 "
      BeginProperty Font 
         Name            =   "나눔고딕"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   630
      TabIndex        =   4
      Top             =   345
      Width           =   3480
   End
End
Attribute VB_Name = "FormSound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim TXT_MAX_LENGTH As Integer

Private Sub txt_Str_Reg_Change(Index As Integer)
    If (LenH(txt_Str_Reg(Index).text) > TXT_MAX_LENGTH) Then
        txt_Str_Reg(Index).text = ""
        'txt_Str_Reg(Index).text = LeftH(txt_Str_Reg(Index).text, TXT_MAX_LENGTH)
    End If
End Sub

Private Sub txt_Str_Guest_Change(Index As Integer)
    If (LenH(txt_Str_Guest(Index).text) > TXT_MAX_LENGTH) Then
        txt_Str_Guest(Index).text = ""
        'txt_Str_Guest(Index).text = LeftH(txt_Str_Guest(Index).text, TXT_MAX_LENGTH)
    End If
End Sub

Private Sub txt_Str_NoRec_Change(Index As Integer)
    If (LenH(txt_Str_NoRec(Index).text) > TXT_MAX_LENGTH) Then
        txt_Str_NoRec(Index).text = ""
        'txt_Str_NoRec(Index).text = LeftH(txt_Str_NoRec(Index).text, TXT_MAX_LENGTH)
    End If
End Sub

Private Sub txt_Str_BKList_Change(Index As Integer)
    If (LenH(txt_Str_BKList(Index).text) > TXT_MAX_LENGTH) Then
        txt_Str_BKList(Index).text = ""
        'txt_Str_BKList(Index).text = LeftH(txt_Str_BKList(Index).text, TXT_MAX_LENGTH)
    End If
End Sub

Private Sub txt_Str_Taxi_Change(Index As Integer)
    If (LenH(txt_Str_Taxi(Index).text) > TXT_MAX_LENGTH) Then
        txt_Str_Taxi(Index).text = ""
        'txt_Str_Taxi(Index).text = LeftH(txt_Str_Taxi(Index).text, TXT_MAX_LENGTH)
    End If
End Sub

Private Sub txt_Str_Day_Change(Index As Integer)
    If (LenH(txt_Str_Day(Index).text) > TXT_MAX_LENGTH) Then
        txt_Str_Day(Index).text = ""
        'txt_Str_Day(Index).text = LeftH(txt_Str_Day(Index).text, TXT_MAX_LENGTH)
    End If
End Sub

Private Sub txt_Str_RegExpDate_Change(Index As Integer)
    If (LenH(txt_Str_RegExpDate(Index).text) > TXT_MAX_LENGTH) Then
        txt_Str_RegExpDate(Index).text = ""
        'txt_Str_RegExpDate(Index).text = LeftH(txt_Str_RegExpDate(Index).text, TXT_MAX_LENGTH)
    End If
End Sub

Private Sub Save_Sound_Config()
    Dim bQryResult As String
    Dim i As Integer
    
    If (chk_SOUND_YN.value = 0) Then
        Glo_SOUND_YN = "N"
        'bQryResult = DataBaseQueryExec(adoConn, "UPDATE tb_config     SET CONTENT = 'N'  WHERE TITLE ='사운드' ", NWERR_GATE_STAY)
    
    Else
        Glo_SOUND_YN = "Y"
        'bQryResult = DataBaseQueryExec(adoConn, "UPDATE tb_config     SET CONTENT = 'Y'  WHERE NAME ='SOUND_YN' ", NWERR_GATE_STAY)

        If (chk_SND_Reg(0).value = 1) Then
            Glo_SND_Lane1_Reg_YN = "Y"
        Else
            Glo_SND_Lane1_Reg_YN = "N"
        End If
        If (chk_SND_Guest(0).value = 1) Then
            Glo_SND_Lane1_Guest_YN = "Y"
            'bQryResult = DataBaseQueryExec(adoConn, "UPDATE tb_config SET CONTENT = 'Y'  WHERE NAME ='Lane1_NoReg' ", NWERR_GATE_STAY)
        Else
            Glo_SND_Lane1_Guest_YN = "N"
            'bQryResult = DataBaseQueryExec(adoConn, "UPDATE tb_config SET CONTENT = 'N'  WHERE NAME ='Lane1_NoReg' ", NWERR_GATE_STAY)
        End If
        If (chk_SND_NoRec(0).value = 1) Then
            Glo_SND_Lane1_NoRec_YN = "Y"
            'bQryResult = DataBaseQueryExec(adoConn, "UPDATE tb_config SET CONTENT = 'Y'  WHERE NAME ='Lane1_NoRec' ", NWERR_GATE_STAY)
        Else
            Glo_SND_Lane1_NoRec_YN = "N"
            'bQryResult = DataBaseQueryExec(adoConn, "UPDATE tb_config SET CONTENT = 'N'  WHERE NAME ='Lane1_NoRec' ", NWERR_GATE_STAY)
        End If
        If (chk_SND_BKList(0).value = 1) Then
            Glo_SND_Lane1_BlackList_YN = "Y"
            'bQryResult = DataBaseQueryExec(adoConn, "UPDATE tb_config SET CONTENT = 'Y'  WHERE NAME ='Lane1_BlackList' ", NWERR_GATE_STAY)
        Else
            Glo_SND_Lane1_BlackList_YN = "N"
            'bQryResult = DataBaseQueryExec(adoConn, "UPDATE tb_config SET CONTENT = 'N'  WHERE NAME ='Lane1_BlackList' ", NWERR_GATE_STAY)
        End If
        If (chk_SND_Taxi(0).value = 1) Then
            Glo_SND_Lane1_Taxi_YN = "Y"
        Else
            Glo_SND_Lane1_Taxi_YN = "N"
        End If
        If (chk_SND_Day(0).value = 1) Then
            Glo_SND_Lane1_Day_YN = "Y"
        Else
            Glo_SND_Lane1_Day_YN = "N"
        End If
        If (chk_SND_RegExpDate(0).value = 1) Then
            Glo_SND_Lane1_RegExpDate_YN = "Y"
        Else
            Glo_SND_Lane1_RegExpDate_YN = "N"
        End If
        If (chk_SND_GuestRegCar(0).value = 1) Then
            Glo_SND_Lane1_GuestRegCar_YN = "Y"
        Else
            Glo_SND_Lane1_GuestRegCar_YN = "N"
        End If
        If (chk_SND_GuestRegCarExpDate(0).value = 1) Then
            Glo_SND_Lane1_GuestRegCarExpDate_YN = "Y"
        Else
            Glo_SND_Lane1_GuestRegCarExpDate_YN = "N"
        End If
        

        If (chk_SND_Reg(1).value = 1) Then
            Glo_SND_Lane2_Reg_YN = "Y"
        Else
            Glo_SND_Lane2_Reg_YN = "N"
        End If
        If (chk_SND_Guest(1).value = 1) Then
            Glo_SND_Lane2_Guest_YN = "Y"
            'bQryResult = DataBaseQueryExec(adoConn, "UPDATE tb_config SET CONTENT = 'Y'  WHERE NAME ='Lane2_NoReg' ", NWERR_GATE_STAY)
        Else
            Glo_SND_Lane2_Guest_YN = "N"
            'bQryResult = DataBaseQueryExec(adoConn, "UPDATE tb_config SET CONTENT = 'N'  WHERE NAME ='Lane2_NoReg' ", NWERR_GATE_STAY)
        End If
        If (chk_SND_NoRec(1).value = 1) Then
            Glo_SND_Lane2_NoRec_YN = "Y"
            'bQryResult = DataBaseQueryExec(adoConn, "UPDATE tb_config SET CONTENT = 'Y'  WHERE NAME ='Lane2_NoRec' ", NWERR_GATE_STAY)
        Else
            Glo_SND_Lane2_NoRec_YN = "N"
            'bQryResult = DataBaseQueryExec(adoConn, "UPDATE tb_config SET CONTENT = 'N'  WHERE NAME ='Lane2_NoRec' ", NWERR_GATE_STAY)
        End If
        If (chk_SND_BKList(1).value = 1) Then
            Glo_SND_Lane2_BlackList_YN = "Y"
            'bQryResult = DataBaseQueryExec(adoConn, "UPDATE tb_config SET CONTENT = 'Y'  WHERE NAME ='Lane2_BlackList' ", NWERR_GATE_STAY)
        Else
            Glo_SND_Lane2_BlackList_YN = "N"
            'bQryResult = DataBaseQueryExec(adoConn, "UPDATE tb_config SET CONTENT = 'N'  WHERE NAME ='Lane2_BlackList' ", NWERR_GATE_STAY)
        End If
        If (chk_SND_Taxi(1).value = 1) Then
            Glo_SND_Lane2_Taxi_YN = "Y"
        Else
            Glo_SND_Lane2_Taxi_YN = "N"
        End If
        If (chk_SND_Day(1).value = 1) Then
            Glo_SND_Lane2_Day_YN = "Y"
        Else
            Glo_SND_Lane2_Day_YN = "N"
        End If
        If (chk_SND_RegExpDate(1).value = 1) Then
            Glo_SND_Lane2_RegExpDate_YN = "Y"
        Else
            Glo_SND_Lane2_RegExpDate_YN = "N"
        End If
        If (chk_SND_GuestRegCar(1).value = 1) Then
            Glo_SND_Lane2_GuestRegCar_YN = "Y"
        Else
            Glo_SND_Lane2_GuestRegCar_YN = "N"
        End If
        If (chk_SND_GuestRegCarExpDate(1).value = 1) Then
            Glo_SND_Lane2_GuestRegCarExpDate_YN = "Y"
        Else
            Glo_SND_Lane2_GuestRegCarExpDate_YN = "N"
        End If
        
        
        If (chk_SND_Reg(2).value = 1) Then
            Glo_SND_Lane3_Reg_YN = "Y"
        Else
            Glo_SND_Lane3_Reg_YN = "N"
        End If
        If (chk_SND_Guest(2).value = 1) Then
            Glo_SND_Lane3_Guest_YN = "Y"
            'bQryResult = DataBaseQueryExec(adoConn, "UPDATE tb_config SET CONTENT = 'Y'  WHERE NAME ='Lane3_NoReg' ", NWERR_GATE_STAY)
        Else
            Glo_SND_Lane3_Guest_YN = "N"
            'bQryResult = DataBaseQueryExec(adoConn, "UPDATE tb_config SET CONTENT = 'N'  WHERE NAME ='Lane3_NoReg' ", NWERR_GATE_STAY)
        End If
        If (chk_SND_NoRec(2).value = 1) Then
            Glo_SND_Lane3_NoRec_YN = "Y"
            'bQryResult = DataBaseQueryExec(adoConn, "UPDATE tb_config SET CONTENT = 'Y'  WHERE NAME ='Lane3_NoRec' ", NWERR_GATE_STAY)
        Else
            Glo_SND_Lane3_NoRec_YN = "N"
            'bQryResult = DataBaseQueryExec(adoConn, "UPDATE tb_config SET CONTENT = 'N'  WHERE NAME ='Lane3_NoRec' ", NWERR_GATE_STAY)
        End If
        If (chk_SND_BKList(2).value = 1) Then
            Glo_SND_Lane3_BlackList_YN = "Y"
            'bQryResult = DataBaseQueryExec(adoConn, "UPDATE tb_config SET CONTENT = 'Y'  WHERE NAME ='Lane3_BlackList' ", NWERR_GATE_STAY)
        Else
            Glo_SND_Lane3_BlackList_YN = "N"
            'bQryResult = DataBaseQueryExec(adoConn, "UPDATE tb_config SET CONTENT = 'N'  WHERE NAME ='Lane3_BlackList' ", NWERR_GATE_STAY)
        End If
        If (chk_SND_Taxi(2).value = 1) Then
            Glo_SND_Lane3_Taxi_YN = "Y"
        Else
            Glo_SND_Lane3_Taxi_YN = "N"
        End If
        If (chk_SND_Day(2).value = 1) Then
            Glo_SND_Lane3_Day_YN = "Y"
        Else
            Glo_SND_Lane3_Day_YN = "N"
        End If
        If (chk_SND_RegExpDate(2).value = 1) Then
            Glo_SND_Lane3_RegExpDate_YN = "Y"
        Else
            Glo_SND_Lane3_RegExpDate_YN = "N"
        End If
        If (chk_SND_GuestRegCar(2).value = 1) Then
            Glo_SND_Lane3_GuestRegCar_YN = "Y"
        Else
            Glo_SND_Lane3_GuestRegCar_YN = "N"
        End If
        If (chk_SND_GuestRegCarExpDate(2).value = 1) Then
            Glo_SND_Lane3_GuestRegCarExpDate_YN = "Y"
        Else
            Glo_SND_Lane3_GuestRegCarExpDate_YN = "N"
        End If
        
        
        If (chk_SND_Reg(3).value = 1) Then
            Glo_SND_Lane4_Reg_YN = "Y"
        Else
            Glo_SND_Lane4_Reg_YN = "N"
        End If
        If (chk_SND_Guest(3).value = 1) Then
            Glo_SND_Lane4_Guest_YN = "Y"
            'bQryResult = DataBaseQueryExec(adoConn, "UPDATE tb_config SET CONTENT = 'Y'  WHERE NAME ='Lane4_NoReg' ", NWERR_GATE_STAY)
        Else
            Glo_SND_Lane4_Guest_YN = "N"
            'bQryResult = DataBaseQueryExec(adoConn, "UPDATE tb_config SET CONTENT = 'N'  WHERE NAME ='Lane4_NoReg' ", NWERR_GATE_STAY)
        End If
        If (chk_SND_NoRec(3).value = 1) Then
            Glo_SND_Lane4_NoRec_YN = "Y"
            'bQryResult = DataBaseQueryExec(adoConn, "UPDATE tb_config SET CONTENT = 'Y'  WHERE NAME ='Lane4_NoRec' ", NWERR_GATE_STAY)
        Else
            Glo_SND_Lane4_NoRec_YN = "N"
            'bQryResult = DataBaseQueryExec(adoConn, "UPDATE tb_config SET CONTENT = 'N'  WHERE NAME ='Lane4_NoRec' ", NWERR_GATE_STAY)
        End If
        If (chk_SND_BKList(3).value = 1) Then
            Glo_SND_Lane4_BlackList_YN = "Y"
            'bQryResult = DataBaseQueryExec(adoConn, "UPDATE tb_config SET CONTENT = 'Y'  WHERE NAME ='Lane4_BlackList' ", NWERR_GATE_STAY)
        Else
            Glo_SND_Lane4_BlackList_YN = "N"
            'bQryResult = DataBaseQueryExec(adoConn, "UPDATE tb_config SET CONTENT = 'N'  WHERE NAME ='Lane4_BlackList' ", NWERR_GATE_STAY)
        End If
        If (chk_SND_Taxi(3).value = 1) Then
            Glo_SND_Lane4_Taxi_YN = "Y"
        Else
            Glo_SND_Lane4_Taxi_YN = "N"
        End If
        If (chk_SND_Day(3).value = 1) Then
            Glo_SND_Lane4_Day_YN = "Y"
        Else
            Glo_SND_Lane4_Day_YN = "N"
        End If
        If (chk_SND_RegExpDate(3).value = 1) Then
            Glo_SND_Lane4_RegExpDate_YN = "Y"
        Else
            Glo_SND_Lane4_RegExpDate_YN = "N"
        End If
        If (chk_SND_GuestRegCar(3).value = 1) Then
            Glo_SND_Lane4_GuestRegCar_YN = "Y"
        Else
            Glo_SND_Lane4_GuestRegCar_YN = "N"
        End If
        If (chk_SND_GuestRegCarExpDate(3).value = 1) Then
            Glo_SND_Lane4_GuestRegCarExpDate_YN = "Y"
        Else
            Glo_SND_Lane4_GuestRegCarExpDate_YN = "N"
        End If
        
        
        If (chk_SND_Reg(4).value = 1) Then
            Glo_SND_Lane5_Reg_YN = "Y"
        Else
            Glo_SND_Lane5_Reg_YN = "N"
        End If
        If (chk_SND_Guest(4).value = 1) Then
            Glo_SND_Lane5_Guest_YN = "Y"
            'bQryResult = DataBaseQueryExec(adoConn, "UPDATE tb_config SET CONTENT = 'Y'  WHERE NAME ='Lane5_NoReg' ", NWERR_GATE_STAY)
        Else
            Glo_SND_Lane5_Guest_YN = "N"
            'bQryResult = DataBaseQueryExec(adoConn, "UPDATE tb_config SET CONTENT = 'N'  WHERE NAME ='Lane5_NoReg' ", NWERR_GATE_STAY)
        End If
        If (chk_SND_NoRec(4).value = 1) Then
            Glo_SND_Lane5_NoRec_YN = "Y"
            'bQryResult = DataBaseQueryExec(adoConn, "UPDATE tb_config SET CONTENT = 'Y'  WHERE NAME ='Lane5_NoRec' ", NWERR_GATE_STAY)
        Else
            Glo_SND_Lane5_NoRec_YN = "N"
            'bQryResult = DataBaseQueryExec(adoConn, "UPDATE tb_config SET CONTENT = 'N'  WHERE NAME ='Lane5_NoRec' ", NWERR_GATE_STAY)
        End If
        If (chk_SND_BKList(4).value = 1) Then
            Glo_SND_Lane5_BlackList_YN = "Y"
            'bQryResult = DataBaseQueryExec(adoConn, "UPDATE tb_config SET CONTENT = 'Y'  WHERE NAME ='Lane5_BlackList' ", NWERR_GATE_STAY)
        Else
            Glo_SND_Lane5_BlackList_YN = "N"
            'bQryResult = DataBaseQueryExec(adoConn, "UPDATE tb_config SET CONTENT = 'N'  WHERE NAME ='Lane5_BlackList' ", NWERR_GATE_STAY)
        End If
        If (chk_SND_Taxi(4).value = 1) Then
            Glo_SND_Lane5_Taxi_YN = "Y"
        Else
            Glo_SND_Lane5_Taxi_YN = "N"
        End If
        If (chk_SND_Day(4).value = 1) Then
            Glo_SND_Lane5_Day_YN = "Y"
        Else
            Glo_SND_Lane5_Day_YN = "N"
        End If
        If (chk_SND_RegExpDate(4).value = 1) Then
            Glo_SND_Lane5_RegExpDate_YN = "Y"
        Else
            Glo_SND_Lane5_RegExpDate_YN = "N"
        End If
        If (chk_SND_GuestRegCar(4).value = 1) Then
            Glo_SND_Lane5_GuestRegCar_YN = "Y"
        Else
            Glo_SND_Lane5_GuestRegCar_YN = "N"
        End If
        If (chk_SND_GuestRegCarExpDate(4).value = 1) Then
            Glo_SND_Lane5_GuestRegCarExpDate_YN = "Y"
        Else
            Glo_SND_Lane5_GuestRegCarExpDate_YN = "N"
        End If
        
        If (chk_SND_Reg(5).value = 1) Then
            Glo_SND_Lane6_Reg_YN = "Y"
        Else
            Glo_SND_Lane6_Reg_YN = "N"
        End If
        If (chk_SND_Guest(5).value = 1) Then
            Glo_SND_Lane6_Guest_YN = "Y"
            'bQryResult = DataBaseQueryExec(adoConn, "UPDATE tb_config SET CONTENT = 'Y'  WHERE NAME ='Lane6_NoReg' ", NWERR_GATE_STAY)
        Else
            Glo_SND_Lane6_Guest_YN = "N"
            'bQryResult = DataBaseQueryExec(adoConn, "UPDATE tb_config SET CONTENT = 'N'  WHERE NAME ='Lane6_NoReg' ", NWERR_GATE_STAY)
        End If
        If (chk_SND_NoRec(5).value = 1) Then
            Glo_SND_Lane6_NoRec_YN = "Y"
            'bQryResult = DataBaseQueryExec(adoConn, "UPDATE tb_config SET CONTENT = 'Y'  WHERE NAME ='Lane6_NoRec' ", NWERR_GATE_STAY)
        Else
            Glo_SND_Lane6_NoRec_YN = "N"
            'bQryResult = DataBaseQueryExec(adoConn, "UPDATE tb_config SET CONTENT = 'N'  WHERE NAME ='Lane6_NoRec' ", NWERR_GATE_STAY)
        End If
        If (chk_SND_BKList(5).value = 1) Then
            Glo_SND_Lane6_BlackList_YN = "Y"
            'bQryResult = DataBaseQueryExec(adoConn, "UPDATE tb_config SET CONTENT = 'Y'  WHERE NAME ='Lane6_BlackList' ", NWERR_GATE_STAY)
        Else
            Glo_SND_Lane6_BlackList_YN = "N"
            'bQryResult = DataBaseQueryExec(adoConn, "UPDATE tb_config SET CONTENT = 'N'  WHERE NAME ='Lane6_BlackList' ", NWERR_GATE_STAY)
        End If
        If (chk_SND_Taxi(5).value = 1) Then
            Glo_SND_Lane6_Taxi_YN = "Y"
        Else
            Glo_SND_Lane6_Taxi_YN = "N"
        End If
        If (chk_SND_Day(5).value = 1) Then
            Glo_SND_Lane6_Day_YN = "Y"
        Else
            Glo_SND_Lane6_Day_YN = "N"
        End If
        If (chk_SND_RegExpDate(5).value = 1) Then
            Glo_SND_Lane6_RegExpDate_YN = "Y"
        Else
            Glo_SND_Lane6_RegExpDate_YN = "N"
        End If
        If (chk_SND_GuestRegCar(5).value = 1) Then
            Glo_SND_Lane6_GuestRegCar_YN = "Y"
        Else
            Glo_SND_Lane6_GuestRegCar_YN = "N"
        End If
        If (chk_SND_GuestRegCarExpDate(5).value = 1) Then
            Glo_SND_Lane6_GuestRegCarExpDate_YN = "Y"
        Else
            Glo_SND_Lane6_GuestRegCarExpDate_YN = "N"
        End If
        
    End If
    
    Call Put_Ini("System Config", "SOUND_YN", Glo_SOUND_YN)
    Call Put_Ini("System Config", "SND_Lane1_Reg_YN", Glo_SND_Lane1_Reg_YN)
    Call Put_Ini("System Config", "SND_Lane1_Guest_YN", Glo_SND_Lane1_Guest_YN)
    Call Put_Ini("System Config", "SND_Lane1_NoRec", Glo_SND_Lane1_NoRec_YN)
    Call Put_Ini("System Config", "SND_Lane1_BlackList_YN", Glo_SND_Lane1_BlackList_YN)
    Call Put_Ini("System Config", "SND_Lane1_Taxi_YN", Glo_SND_Lane1_Taxi_YN)
    Call Put_Ini("System Config", "SND_Lane1_Day_YN", Glo_SND_Lane1_Day_YN)
    Call Put_Ini("System Config", "SND_Lane1_RegExpDate_YN", Glo_SND_Lane1_RegExpDate_YN)
    Call Put_Ini("System Config", "SND_Lane1_GuestRegCar_YN", Glo_SND_Lane1_GuestRegCar_YN)
    Call Put_Ini("System Config", "SND_Lane1_GuestRegCarExpDate_YN", Glo_SND_Lane1_GuestRegCarExpDate_YN)
    Call Put_Ini("System Config", "SND_Lane2_Reg_YN", Glo_SND_Lane2_Reg_YN)
    Call Put_Ini("System Config", "SND_Lane2_Guest_YN", Glo_SND_Lane2_Guest_YN)
    Call Put_Ini("System Config", "SND_Lane2_NoRec_YN", Glo_SND_Lane2_NoRec_YN)
    Call Put_Ini("System Config", "SND_Lane2_BlackList_YN", Glo_SND_Lane2_BlackList_YN)
    Call Put_Ini("System Config", "SND_Lane2_Taxi_YN", Glo_SND_Lane2_Taxi_YN)
    Call Put_Ini("System Config", "SND_Lane2_Day_YN", Glo_SND_Lane2_Day_YN)
    Call Put_Ini("System Config", "SND_Lane2_RegExpDate_YN", Glo_SND_Lane2_RegExpDate_YN)
    Call Put_Ini("System Config", "SND_Lane2_GuestRegCar_YN", Glo_SND_Lane2_GuestRegCar_YN)
    Call Put_Ini("System Config", "SND_Lane2_GuestRegCarExpDate_YN", Glo_SND_Lane2_GuestRegCarExpDate_YN)
    Call Put_Ini("System Config", "SND_Lane3_Reg_YN", Glo_SND_Lane3_Reg_YN)
    Call Put_Ini("System Config", "SND_Lane3_Guest_YN", Glo_SND_Lane3_Guest_YN)
    Call Put_Ini("System Config", "SND_Lane3_NoRec_YN", Glo_SND_Lane3_NoRec_YN)
    Call Put_Ini("System Config", "SND_Lane3_BlackList_YN", Glo_SND_Lane3_BlackList_YN)
    Call Put_Ini("System Config", "SND_Lane3_Taxi_YN", Glo_SND_Lane3_Taxi_YN)
    Call Put_Ini("System Config", "SND_Lane3_Day_YN", Glo_SND_Lane3_Day_YN)
    Call Put_Ini("System Config", "SND_Lane3_RegExpDate_YN", Glo_SND_Lane3_RegExpDate_YN)
    Call Put_Ini("System Config", "SND_Lane3_GuestRegCar_YN", Glo_SND_Lane3_GuestRegCar_YN)
    Call Put_Ini("System Config", "SND_Lane3_GuestRegCarExpDate_YN", Glo_SND_Lane3_GuestRegCarExpDate_YN)
    Call Put_Ini("System Config", "SND_Lane4_Reg_YN", Glo_SND_Lane4_Reg_YN)
    Call Put_Ini("System Config", "SND_Lane4_Guest_YN", Glo_SND_Lane4_Guest_YN)
    Call Put_Ini("System Config", "SND_Lane4_NoRec_YN", Glo_SND_Lane4_NoRec_YN)
    Call Put_Ini("System Config", "SND_Lane4_BlackList_YN", Glo_SND_Lane4_BlackList_YN)
    Call Put_Ini("System Config", "SND_Lane4_Taxi_YN", Glo_SND_Lane4_Taxi_YN)
    Call Put_Ini("System Config", "SND_Lane4_Day_YN", Glo_SND_Lane4_Day_YN)
    Call Put_Ini("System Config", "SND_Lane4_RegExpDate_YN", Glo_SND_Lane4_RegExpDate_YN)
    Call Put_Ini("System Config", "SND_Lane4_GuestRegCar_YN", Glo_SND_Lane4_GuestRegCar_YN)
    Call Put_Ini("System Config", "SND_Lane4_GuestRegCarExpDate_YN", Glo_SND_Lane4_GuestRegCarExpDate_YN)
    Call Put_Ini("System Config", "SND_Lane5_Reg_YN", Glo_SND_Lane5_Reg_YN)
    Call Put_Ini("System Config", "SND_Lane5_Guest_YN", Glo_SND_Lane5_Guest_YN)
    Call Put_Ini("System Config", "SND_Lane5_NoRec_YN", Glo_SND_Lane5_NoRec_YN)
    Call Put_Ini("System Config", "SND_Lane5_BlackList_YN", Glo_SND_Lane5_BlackList_YN)
    Call Put_Ini("System Config", "SND_Lane5_Taxi_YN", Glo_SND_Lane5_Taxi_YN)
    Call Put_Ini("System Config", "SND_Lane5_Day_YN", Glo_SND_Lane5_Day_YN)
    Call Put_Ini("System Config", "SND_Lane5_RegExpDate_YN", Glo_SND_Lane5_RegExpDate_YN)
    Call Put_Ini("System Config", "SND_Lane5_GuestRegCar_YN", Glo_SND_Lane5_GuestRegCar_YN)
    Call Put_Ini("System Config", "SND_Lane5_GuestRegCarExpDate_YN", Glo_SND_Lane5_GuestRegCarExpDate_YN)
    Call Put_Ini("System Config", "SND_Lane6_Reg_YN", Glo_SND_Lane6_Reg_YN)
    Call Put_Ini("System Config", "SND_Lane6_Guest_YN", Glo_SND_Lane6_Guest_YN)
    Call Put_Ini("System Config", "SND_Lane6_NoRec_YN", Glo_SND_Lane6_NoRec_YN)
    Call Put_Ini("System Config", "SND_Lane6_BlackList_YN", Glo_SND_Lane6_BlackList_YN)
    Call Put_Ini("System Config", "SND_Lane6_Taxi_YN", Glo_SND_Lane6_Taxi_YN)
    Call Put_Ini("System Config", "SND_Lane6_Day_YN", Glo_SND_Lane6_Day_YN)
    Call Put_Ini("System Config", "SND_Lane6_RegExpDate_YN", Glo_SND_Lane6_RegExpDate_YN)
    Call Put_Ini("System Config", "SND_Lane6_GuestRegCar_YN", Glo_SND_Lane6_GuestRegCar_YN)
    Call Put_Ini("System Config", "SND_Lane6_GuestRegCarExpDate_YN", Glo_SND_Lane6_GuestRegCarExpDate_YN)



    For i = 0 To MAX_LANE_COUNT - 1
        Glo_SNDFILE_Reg(i) = Tmp_SNDFILE_Reg(i)
        Glo_SNDFILE_Guest(i) = Tmp_SNDFILE_Guest(i)
        Glo_SNDFILE_NoRec(i) = Tmp_SNDFILE_NoRec(i)
        Glo_SNDFILE_BlackList(i) = Tmp_SNDFILE_BlackList(i)
        Glo_SNDFILE_Taxi(i) = Tmp_SNDFILE_Taxi(i)
        Glo_SNDFILE_Day(i) = Tmp_SNDFILE_Day(i)
        Glo_SNDFILE_RegExpDate(i) = Tmp_SNDFILE_RegExpDate(i)
        Glo_SNDFILE_GuestRegCar(i) = Tmp_SNDFILE_GuestRegCar(i)
        Glo_SNDFILE_GuestRegCarExpDate(i) = Tmp_SNDFILE_GuestRegCarExpDate(i)
    Next i

    Call Put_Ini("System Config", "SNDFILE_Lane1_Reg", Glo_SNDFILE_Reg(0))
    Call Put_Ini("System Config", "SNDFILE_Lane1_Guest", Glo_SNDFILE_Guest(0))
    Call Put_Ini("System Config", "SNDFILE_Lane1_NoRec", Glo_SNDFILE_NoRec(0))
    Call Put_Ini("System Config", "SNDFILE_Lane1_BlackList", Glo_SNDFILE_BlackList(0))
    Call Put_Ini("System Config", "SNDFILE_Lane1_Taxi", Glo_SNDFILE_Taxi(0))
    Call Put_Ini("System Config", "SNDFILE_Lane1_Day", Glo_SNDFILE_Day(0))
    Call Put_Ini("System Config", "SNDFILE_Lane1_RegExpDate", Glo_SNDFILE_RegExpDate(0))
    Call Put_Ini("System Config", "SNDFILE_Lane1_GuestRegCar", Glo_SNDFILE_GuestRegCar(0))
    Call Put_Ini("System Config", "SNDFILE_Lane1_GuestRegCarExpDate", Glo_SNDFILE_GuestRegCarExpDate(0))
    
    Call Put_Ini("System Config", "SNDFILE_Lane2_Reg", Glo_SNDFILE_Reg(1))
    Call Put_Ini("System Config", "SNDFILE_Lane2_Guest", Glo_SNDFILE_Guest(1))
    Call Put_Ini("System Config", "SNDFILE_Lane2_NoRec", Glo_SNDFILE_NoRec(1))
    Call Put_Ini("System Config", "SNDFILE_Lane2_BlackList", Glo_SNDFILE_BlackList(1))
    Call Put_Ini("System Config", "SNDFILE_Lane2_Taxi", Glo_SNDFILE_Taxi(1))
    Call Put_Ini("System Config", "SNDFILE_Lane2_Day", Glo_SNDFILE_Day(1))
    Call Put_Ini("System Config", "SNDFILE_Lane2_RegExpDate", Glo_SNDFILE_RegExpDate(1))
    Call Put_Ini("System Config", "SNDFILE_Lane2_GuestRegCar", Glo_SNDFILE_GuestRegCar(1))
    Call Put_Ini("System Config", "SNDFILE_Lane2_GuestRegCarExpDate", Glo_SNDFILE_GuestRegCarExpDate(1))
    
    Call Put_Ini("System Config", "SNDFILE_Lane3_Reg", Glo_SNDFILE_Reg(2))
    Call Put_Ini("System Config", "SNDFILE_Lane3_Guest", Glo_SNDFILE_Guest(2))
    Call Put_Ini("System Config", "SNDFILE_Lane3_NoRec", Glo_SNDFILE_NoRec(2))
    Call Put_Ini("System Config", "SNDFILE_Lane3_BlackList", Glo_SNDFILE_BlackList(2))
    Call Put_Ini("System Config", "SNDFILE_Lane3_Taxi", Glo_SNDFILE_Taxi(2))
    Call Put_Ini("System Config", "SNDFILE_Lane3_Day", Glo_SNDFILE_Day(2))
    Call Put_Ini("System Config", "SNDFILE_Lane3_RegExpDate", Glo_SNDFILE_RegExpDate(2))
    Call Put_Ini("System Config", "SNDFILE_Lane3_GuestRegCar", Glo_SNDFILE_GuestRegCar(2))
    Call Put_Ini("System Config", "SNDFILE_Lane3_GuestRegCarExpDate", Glo_SNDFILE_GuestRegCarExpDate(2))
    
    Call Put_Ini("System Config", "SNDFILE_Lane4_Reg", Glo_SNDFILE_Reg(3))
    Call Put_Ini("System Config", "SNDFILE_Lane4_Guest", Glo_SNDFILE_Guest(3))
    Call Put_Ini("System Config", "SNDFILE_Lane4_NoRec", Glo_SNDFILE_NoRec(3))
    Call Put_Ini("System Config", "SNDFILE_Lane4_BlackList", Glo_SNDFILE_BlackList(3))
    Call Put_Ini("System Config", "SNDFILE_Lane4_Taxi", Glo_SNDFILE_Taxi(3))
    Call Put_Ini("System Config", "SNDFILE_Lane4_Day", Glo_SNDFILE_Day(3))
    Call Put_Ini("System Config", "SNDFILE_Lane4_RegExpDate", Glo_SNDFILE_RegExpDate(3))
    Call Put_Ini("System Config", "SNDFILE_Lane4_GuestRegCar", Glo_SNDFILE_GuestRegCar(3))
    Call Put_Ini("System Config", "SNDFILE_Lane4_GuestRegCarExpDate", Glo_SNDFILE_GuestRegCarExpDate(3))
    
    Call Put_Ini("System Config", "SNDFILE_Lane5_Reg", Glo_SNDFILE_Reg(4))
    Call Put_Ini("System Config", "SNDFILE_Lane5_Guest", Glo_SNDFILE_Guest(4))
    Call Put_Ini("System Config", "SNDFILE_Lane5_NoRec", Glo_SNDFILE_NoRec(4))
    Call Put_Ini("System Config", "SNDFILE_Lane5_BlackList", Glo_SNDFILE_BlackList(4))
    Call Put_Ini("System Config", "SNDFILE_Lane5_Taxi", Glo_SNDFILE_Taxi(4))
    Call Put_Ini("System Config", "SNDFILE_Lane5_Day", Glo_SNDFILE_Day(4))
    Call Put_Ini("System Config", "SNDFILE_Lane5_RegExpDate", Glo_SNDFILE_RegExpDate(4))
    Call Put_Ini("System Config", "SNDFILE_Lane5_GuestRegCar", Glo_SNDFILE_GuestRegCar(4))
    Call Put_Ini("System Config", "SNDFILE_Lane5_GuestRegCarExpDate", Glo_SNDFILE_GuestRegCarExpDate(4))
    
    Call Put_Ini("System Config", "SNDFILE_Lane6_Reg", Glo_SNDFILE_Reg(5))
    Call Put_Ini("System Config", "SNDFILE_Lane6_Guest", Glo_SNDFILE_Guest(5))
    Call Put_Ini("System Config", "SNDFILE_Lane6_NoRec", Glo_SNDFILE_NoRec(5))
    Call Put_Ini("System Config", "SNDFILE_Lane6_BlackList", Glo_SNDFILE_BlackList(5))
    Call Put_Ini("System Config", "SNDFILE_Lane6_Taxi", Glo_SNDFILE_Taxi(5))
    Call Put_Ini("System Config", "SNDFILE_Lane6_Day", Glo_SNDFILE_Day(5))
    Call Put_Ini("System Config", "SNDFILE_Lane6_RegExpDate", Glo_SNDFILE_RegExpDate(5))
    Call Put_Ini("System Config", "SNDFILE_Lane6_GuestRegCar", Glo_SNDFILE_GuestRegCar(5))
    Call Put_Ini("System Config", "SNDFILE_Lane6_GuestRegCarExpDate", Glo_SNDFILE_GuestRegCarExpDate(5))
End Sub


Private Sub Load_Sound_Config()
'    Call DB_CFG_Init("사운드")

    ' 사운드 파일명 로드
    Dim i As Integer
    For i = 0 To MAX_LANE_COUNT - 1
        Tmp_SNDFILE_Reg(i) = Glo_SNDFILE_Reg(i)
        Tmp_SNDFILE_Guest(i) = Glo_SNDFILE_Guest(i)
        Tmp_SNDFILE_NoRec(i) = Glo_SNDFILE_NoRec(i)
        Tmp_SNDFILE_BlackList(i) = Glo_SNDFILE_BlackList(i)
        Tmp_SNDFILE_Taxi(i) = Glo_SNDFILE_Taxi(i)
        Tmp_SNDFILE_Day(i) = Glo_SNDFILE_Day(i)
        Tmp_SNDFILE_RegExpDate(i) = Glo_SNDFILE_RegExpDate(i)
        Tmp_SNDFILE_GuestRegCar(i) = Glo_SNDFILE_GuestRegCar(i)
        Tmp_SNDFILE_GuestRegCarExpDate(i) = Glo_SNDFILE_GuestRegCar(i)
    Next i
    

    If (Glo_SOUND_YN = "Y") Then
        'frmSOUND.Enabled = True
        chk_SOUND_YN.value = 1
        For i = 0 To MAX_LANE_COUNT - 1
            chk_SND_Reg(i).Enabled = True
            chk_SND_Guest(i).Enabled = True
            chk_SND_NoRec(i).Enabled = True
            chk_SND_BKList(i).Enabled = True
            chk_SND_Taxi(i).Enabled = True
            chk_SND_Day(i).Enabled = True
            chk_SND_RegExpDate(i).Enabled = True
            chk_SND_GuestRegCar(i).Enabled = True
            chk_SND_GuestRegCarExpDate(i).Enabled = True
        Next i
    Else
        'frmSOUND.Enabled = False
        chk_SOUND_YN.value = 0
        For i = 0 To MAX_LANE_COUNT - 1
            chk_SND_Reg(i).Enabled = False
            chk_SND_Guest(i).Enabled = False
            chk_SND_NoRec(i).Enabled = False
            chk_SND_BKList(i).Enabled = False
            chk_SND_Taxi(i).Enabled = False
            chk_SND_Day(i).Enabled = False
            chk_SND_RegExpDate(i).Enabled = False
            chk_SND_GuestRegCar(i).Enabled = False
            chk_SND_GuestRegCarExpDate(i).Enabled = False
        Next i
    End If
    
    
    
    If (Glo_SND_Lane1_Reg_YN = "Y") Then
        chk_SND_Reg(0).value = 1
    Else
        chk_SND_Reg(0).value = 0
    End If
    If (Glo_SND_Lane1_Guest_YN = "Y") Then
        chk_SND_Guest(0).value = 1
    Else
        chk_SND_Guest(0).value = 0
    End If
    If (Glo_SND_Lane1_NoRec_YN = "Y") Then
        chk_SND_NoRec(0).value = 1
    Else
        chk_SND_NoRec(0).value = 0
    End If
    If (Glo_SND_Lane1_BlackList_YN = "Y") Then
        chk_SND_BKList(0).value = 1
    Else
        chk_SND_BKList(0).value = 0
    End If
    If (Glo_SND_Lane1_Taxi_YN = "Y") Then
        chk_SND_Taxi(0).value = 1
    Else
        chk_SND_Taxi(0).value = 0
    End If
    If (Glo_SND_Lane1_Day_YN = "Y") Then
        chk_SND_Day(0).value = 1
    Else
        chk_SND_Day(0).value = 0
    End If
    If (Glo_SND_Lane1_RegExpDate_YN = "Y") Then
        chk_SND_RegExpDate(0).value = 1
    Else
        chk_SND_RegExpDate(0).value = 0
    End If
    If (Glo_SND_Lane1_GuestRegCar_YN = "Y") Then
        chk_SND_GuestRegCar(0).value = 1
    Else
        chk_SND_GuestRegCar(0).value = 0
    End If
    If (Glo_SND_Lane1_GuestRegCarExpDate_YN = "Y") Then
        chk_SND_GuestRegCarExpDate(0).value = 1
    Else
        chk_SND_GuestRegCarExpDate(0).value = 0
    End If
'''''''''''''''''''''''''''''''''''''
    If (Glo_SND_Lane2_Reg_YN = "Y") Then
        chk_SND_Reg(1).value = 1
    Else
        chk_SND_Reg(1).value = 0
    End If
    If (Glo_SND_Lane2_Guest_YN = "Y") Then
        chk_SND_Guest(1).value = 1
    Else
        chk_SND_Guest(1).value = 0
    End If
    If (Glo_SND_Lane2_NoRec_YN = "Y") Then
        chk_SND_NoRec(1).value = 1
    Else
        chk_SND_NoRec(1).value = 0
    End If
    If (Glo_SND_Lane2_BlackList_YN = "Y") Then
        chk_SND_BKList(1).value = 1
    Else
        chk_SND_BKList(1).value = 0
    End If
    If (Glo_SND_Lane2_Taxi_YN = "Y") Then
        chk_SND_Taxi(1).value = 1
    Else
        chk_SND_Taxi(1).value = 0
    End If
    If (Glo_SND_Lane2_Day_YN = "Y") Then
        chk_SND_Day(1).value = 1
    Else
        chk_SND_Day(1).value = 0
    End If
    If (Glo_SND_Lane2_RegExpDate_YN = "Y") Then
        chk_SND_RegExpDate(1).value = 1
    Else
        chk_SND_RegExpDate(1).value = 0
    End If
    If (Glo_SND_Lane2_GuestRegCar_YN = "Y") Then
        chk_SND_GuestRegCar(1).value = 1
    Else
        chk_SND_GuestRegCar(1).value = 0
    End If
    If (Glo_SND_Lane2_GuestRegCarExpDate_YN = "Y") Then
        chk_SND_GuestRegCarExpDate(1).value = 1
    Else
        chk_SND_GuestRegCarExpDate(1).value = 0
    End If
    ''''''''''''''''''''''''''''''''''''''''''''
    If (Glo_SND_Lane3_Reg_YN = "Y") Then
        chk_SND_Reg(2).value = 1
    Else
        chk_SND_Reg(2).value = 0
    End If
    If (Glo_SND_Lane3_Guest_YN = "Y") Then
        chk_SND_Guest(2).value = 1
    Else
        chk_SND_Guest(2).value = 0
    End If
    If (Glo_SND_Lane3_NoRec_YN = "Y") Then
        chk_SND_NoRec(2).value = 1
    Else
        chk_SND_NoRec(2).value = 0
    End If
    If (Glo_SND_Lane3_BlackList_YN = "Y") Then
        chk_SND_BKList(2).value = 1
    Else
        chk_SND_BKList(2).value = 0
    End If
    If (Glo_SND_Lane3_Taxi_YN = "Y") Then
        chk_SND_Taxi(2).value = 1
    Else
        chk_SND_Taxi(2).value = 0
    End If
    If (Glo_SND_Lane3_Day_YN = "Y") Then
        chk_SND_Day(2).value = 1
    Else
        chk_SND_Day(2).value = 0
    End If
    If (Glo_SND_Lane3_RegExpDate_YN = "Y") Then
        chk_SND_RegExpDate(2).value = 1
    Else
        chk_SND_RegExpDate(2).value = 0
    End If
    If (Glo_SND_Lane3_GuestRegCar_YN = "Y") Then
        chk_SND_GuestRegCar(2).value = 1
    Else
        chk_SND_GuestRegCar(2).value = 0
    End If
    If (Glo_SND_Lane3_GuestRegCarExpDate_YN = "Y") Then
        chk_SND_GuestRegCarExpDate(2).value = 1
    Else
        chk_SND_GuestRegCarExpDate(2).value = 0
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''

    If (Glo_SND_Lane4_Reg_YN = "Y") Then
        chk_SND_Reg(3).value = 1
    Else
        chk_SND_Reg(3).value = 0
    End If
    If (Glo_SND_Lane4_Guest_YN = "Y") Then
        chk_SND_Guest(3).value = 1
    Else
        chk_SND_Guest(3).value = 0
    End If
    If (Glo_SND_Lane4_NoRec_YN = "Y") Then
        chk_SND_NoRec(3).value = 1
    Else
        chk_SND_NoRec(3).value = 0
    End If
    If (Glo_SND_Lane4_BlackList_YN = "Y") Then
        chk_SND_BKList(3).value = 1
    Else
        chk_SND_BKList(3).value = 0
    End If
    If (Glo_SND_Lane4_Taxi_YN = "Y") Then
        chk_SND_Taxi(3).value = 1
    Else
        chk_SND_Taxi(3).value = 0
    End If
    If (Glo_SND_Lane4_Day_YN = "Y") Then
        chk_SND_Day(3).value = 1
    Else
        chk_SND_Day(3).value = 0
    End If
    If (Glo_SND_Lane4_RegExpDate_YN = "Y") Then
        chk_SND_RegExpDate(3).value = 1
    Else
        chk_SND_RegExpDate(3).value = 0
    End If
    If (Glo_SND_Lane4_GuestRegCar_YN = "Y") Then
        chk_SND_GuestRegCar(3).value = 1
    Else
        chk_SND_GuestRegCar(3).value = 0
    End If
    If (Glo_SND_Lane4_GuestRegCarExpDate_YN = "Y") Then
        chk_SND_GuestRegCarExpDate(3).value = 1
    Else
        chk_SND_GuestRegCarExpDate(3).value = 0
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''

    If (Glo_SND_Lane5_Reg_YN = "Y") Then
        chk_SND_Reg(4).value = 1
    Else
        chk_SND_Reg(4).value = 0
    End If
    If (Glo_SND_Lane5_Guest_YN = "Y") Then
        chk_SND_Guest(4).value = 1
    Else
        chk_SND_Guest(4).value = 0
    End If
    If (Glo_SND_Lane5_NoRec_YN = "Y") Then
        chk_SND_NoRec(4).value = 1
    Else
        chk_SND_NoRec(4).value = 0
    End If
    If (Glo_SND_Lane5_BlackList_YN = "Y") Then
        chk_SND_BKList(4).value = 1
    Else
        chk_SND_BKList(4).value = 0
    End If
    If (Glo_SND_Lane5_Taxi_YN = "Y") Then
        chk_SND_Taxi(4).value = 1
    Else
        chk_SND_Taxi(4).value = 0
    End If
    If (Glo_SND_Lane5_Day_YN = "Y") Then
        chk_SND_Day(4).value = 1
    Else
        chk_SND_Day(4).value = 0
    End If
    If (Glo_SND_Lane5_RegExpDate_YN = "Y") Then
        chk_SND_RegExpDate(4).value = 1
    Else
        chk_SND_RegExpDate(4).value = 0
    End If
    If (Glo_SND_Lane5_GuestRegCar_YN = "Y") Then
        chk_SND_GuestRegCar(4).value = 1
    Else
        chk_SND_GuestRegCar(4).value = 0
    End If
    If (Glo_SND_Lane5_GuestRegCarExpDate_YN = "Y") Then
        chk_SND_GuestRegCarExpDate(4).value = 1
    Else
        chk_SND_GuestRegCarExpDate(4).value = 0
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''
    
    If (Glo_SND_Lane6_Reg_YN = "Y") Then
        chk_SND_Reg(5).value = 1
    Else
        chk_SND_Reg(5).value = 0
    End If
    If (Glo_SND_Lane6_Guest_YN = "Y") Then
        chk_SND_Guest(5).value = 1
    Else
        chk_SND_Guest(5).value = 0
    End If
    If (Glo_SND_Lane6_NoRec_YN = "Y") Then
        chk_SND_NoRec(5).value = 1
    Else
        chk_SND_NoRec(5).value = 0
    End If
    If (Glo_SND_Lane6_BlackList_YN = "Y") Then
        chk_SND_BKList(5).value = 1
    Else
        chk_SND_BKList(5).value = 0
    End If
    If (Glo_SND_Lane6_Taxi_YN = "Y") Then
        chk_SND_Taxi(5).value = 1
    Else
        chk_SND_Taxi(5).value = 0
    End If
    If (Glo_SND_Lane6_Day_YN = "Y") Then
        chk_SND_Day(5).value = 1
    Else
        chk_SND_Day(5).value = 0
    End If
    If (Glo_SND_Lane6_RegExpDate_YN = "Y") Then
        chk_SND_RegExpDate(5).value = 1
    Else
        chk_SND_RegExpDate(5).value = 0
    End If
    If (Glo_SND_Lane6_GuestRegCar_YN = "Y") Then
        chk_SND_GuestRegCar(5).value = 1
    Else
        chk_SND_GuestRegCar(5).value = 0
    End If
    If (Glo_SND_Lane6_GuestRegCarExpDate_YN = "Y") Then
        chk_SND_GuestRegCarExpDate(5).value = 1
    Else
        chk_SND_GuestRegCarExpDate(5).value = 0
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''
End Sub


Private Sub Load_Disp_Config()
    Dim rs As ADODB.Recordset
    Dim i As Integer
    
    If (Glo_Display = "전광판" Or Glo_Display = "전광판(풀컬러)") Then
        For i = 0 To MAX_LANE_COUNT - 1
            cmb_Disp1EmgColorReg(i).Visible = False
            cmb_Disp2EmgColorReg(i).Visible = False
            cmb_Disp1EmgColorGuest(i).Visible = False
            cmb_Disp2EmgColorGuest(i).Visible = False
            cmb_Disp1EmgColorNoRec(i).Visible = False
            cmb_Disp2EmgColorNoRec(i).Visible = False
            cmb_Disp1EmgColorBKList(i).Visible = False
            cmb_Disp2EmgColorBKList(i).Visible = False
            cmb_Disp1EmgColorTaxi(i).Visible = False
            cmb_Disp2EmgColorTaxi(i).Visible = False
            cmb_Disp1EmgColorDay(i).Visible = False
            cmb_Disp2EmgColorDay(i).Visible = False
            cmb_Disp1EmgColorRegExpDate(i).Visible = False
            cmb_Disp2EmgColorRegExpDate(i).Visible = False
            cmb_Disp1EmgColorGuestRegCar(i).Visible = False
            cmb_Disp2EmgColorGuestRegCar(i).Visible = False
            cmb_Disp1EmgColorGuestRegCarExpDate(i).Visible = False
            cmb_Disp2EmgColorGuestRegCarExpDate(i).Visible = False
        Next i
    ElseIf (Glo_Display = "전광판(풀컬러)_FW7") Then
    
        For i = 0 To MAX_LANE_COUNT - 1
            cmb_Disp1EmgColorReg(i).Visible = True
            cmb_Disp2EmgColorReg(i).Visible = True
            cmb_Disp1EmgColorGuest(i).Visible = True
            cmb_Disp2EmgColorGuest(i).Visible = True
            cmb_Disp1EmgColorNoRec(i).Visible = True
            cmb_Disp2EmgColorNoRec(i).Visible = True
            cmb_Disp1EmgColorBKList(i).Visible = True
            cmb_Disp2EmgColorBKList(i).Visible = True
            cmb_Disp1EmgColorTaxi(i).Visible = True
            cmb_Disp2EmgColorTaxi(i).Visible = True
            cmb_Disp1EmgColorDay(i).Visible = True
            cmb_Disp2EmgColorDay(i).Visible = True
            cmb_Disp1EmgColorRegExpDate(i).Visible = True
            cmb_Disp2EmgColorRegExpDate(i).Visible = True
            cmb_Disp1EmgColorGuestRegCar(i).Visible = True
            cmb_Disp2EmgColorGuestRegCar(i).Visible = True
            cmb_Disp1EmgColorGuestRegCarExpDate(i).Visible = True
            cmb_Disp2EmgColorGuestRegCarExpDate(i).Visible = True
        Next i
        
        
        Set rs = New ADODB.Recordset
        rs.Open "SELECT * FROM tb_config ", adoConn
    
        Do While Not (rs.EOF)
    
            If (rs!name = "LANE1_Disp1EmgColorReg") Then
                cmb_Disp1EmgColorReg(0).text = rs!Content:
            ElseIf (rs!name = "LANE1_Disp2EmgColorReg") Then cmb_Disp2EmgColorReg(0).text = rs!Content
            ElseIf (rs!name = "LANE1_Disp1EmgColorGuest") Then cmb_Disp1EmgColorGuest(0).text = rs!Content
            ElseIf (rs!name = "LANE1_Disp2EmgColorGuest") Then cmb_Disp2EmgColorGuest(0).text = rs!Content
            ElseIf (rs!name = "LANE1_Disp1EmgColorNoRec") Then cmb_Disp1EmgColorNoRec(0).text = rs!Content
            ElseIf (rs!name = "LANE1_Disp2EmgColorNoRec") Then cmb_Disp2EmgColorNoRec(0).text = rs!Content
            ElseIf (rs!name = "LANE1_Disp1EmgColorBKList") Then cmb_Disp1EmgColorBKList(0).text = rs!Content
            ElseIf (rs!name = "LANE1_Disp2EmgColorBKList") Then cmb_Disp2EmgColorBKList(0).text = rs!Content
            ElseIf (rs!name = "LANE1_Disp1EmgColorTaxi") Then cmb_Disp1EmgColorTaxi(0).text = rs!Content
            ElseIf (rs!name = "LANE1_Disp2EmgColorTaxi") Then cmb_Disp2EmgColorTaxi(0).text = rs!Content
            ElseIf (rs!name = "LANE1_Disp1EmgColorDay") Then cmb_Disp1EmgColorDay(0).text = rs!Content
            ElseIf (rs!name = "LANE1_Disp2EmgColorDay") Then cmb_Disp2EmgColorDay(0).text = rs!Content
            ElseIf (rs!name = "LANE1_Disp1EmgColorRegExpDate") Then cmb_Disp1EmgColorRegExpDate(0).text = rs!Content
            ElseIf (rs!name = "LANE1_Disp2EmgColorRegExpDate") Then cmb_Disp2EmgColorRegExpDate(0).text = rs!Content
            ElseIf (rs!name = "LANE1_Disp1EmgColorGuestRegCar") Then cmb_Disp1EmgColorGuestRegCar(0).text = rs!Content
            ElseIf (rs!name = "LANE1_Disp2EmgColorGuestRegCar") Then cmb_Disp2EmgColorGuestRegCar(0).text = rs!Content
            ElseIf (rs!name = "LANE1_Disp1EmgColorGuestRegCarExpDate") Then cmb_Disp1EmgColorGuestRegCarExpDate(0).text = rs!Content
            ElseIf (rs!name = "LANE1_Disp2EmgColorGuestRegCarExpDate") Then cmb_Disp2EmgColorGuestRegCarExpDate(0).text = rs!Content
            
            ElseIf (rs!name = "LANE2_Disp1EmgColorReg") Then cmb_Disp1EmgColorReg(1).text = rs!Content
            ElseIf (rs!name = "LANE2_Disp2EmgColorReg") Then cmb_Disp2EmgColorReg(1).text = rs!Content
            ElseIf (rs!name = "LANE2_Disp1EmgColorGuest") Then cmb_Disp1EmgColorGuest(1).text = rs!Content
            ElseIf (rs!name = "LANE2_Disp2EmgColorGuest") Then cmb_Disp2EmgColorGuest(1).text = rs!Content
            ElseIf (rs!name = "LANE2_Disp1EmgColorNoRec") Then cmb_Disp1EmgColorNoRec(1).text = rs!Content
            ElseIf (rs!name = "LANE2_Disp2EmgColorNoRec") Then cmb_Disp2EmgColorNoRec(1).text = rs!Content
            ElseIf (rs!name = "LANE2_Disp1EmgColorBKList") Then cmb_Disp1EmgColorBKList(1).text = rs!Content
            ElseIf (rs!name = "LANE2_Disp2EmgColorBKList") Then cmb_Disp2EmgColorBKList(1).text = rs!Content
            ElseIf (rs!name = "LANE2_Disp1EmgColorTaxi") Then cmb_Disp1EmgColorTaxi(1).text = rs!Content
            ElseIf (rs!name = "LANE2_Disp2EmgColorTaxi") Then cmb_Disp2EmgColorTaxi(1).text = rs!Content
            ElseIf (rs!name = "LANE2_Disp1EmgColorDay") Then cmb_Disp1EmgColorDay(1).text = rs!Content
            ElseIf (rs!name = "LANE2_Disp2EmgColorDay") Then cmb_Disp2EmgColorDay(1).text = rs!Content
            ElseIf (rs!name = "LANE2_Disp1EmgColorRegExpDate") Then cmb_Disp1EmgColorRegExpDate(1).text = rs!Content
            ElseIf (rs!name = "LANE2_Disp2EmgColorRegExpDate") Then cmb_Disp2EmgColorRegExpDate(1).text = rs!Content
            ElseIf (rs!name = "LANE2_Disp1EmgColorGuestRegCar") Then cmb_Disp1EmgColorGuestRegCar(1).text = rs!Content
            ElseIf (rs!name = "LANE2_Disp2EmgColorGuestRegCar") Then cmb_Disp2EmgColorGuestRegCar(1).text = rs!Content
            ElseIf (rs!name = "LANE2_Disp1EmgColorGuestRegCarExpDate") Then cmb_Disp1EmgColorGuestRegCarExpDate(1).text = rs!Content
            ElseIf (rs!name = "LANE2_Disp2EmgColorGuestRegCarExpDate") Then cmb_Disp2EmgColorGuestRegCarExpDate(1).text = rs!Content
            
            ElseIf (rs!name = "LANE3_Disp1EmgColorReg") Then cmb_Disp1EmgColorReg(2).text = rs!Content
            ElseIf (rs!name = "LANE3_Disp2EmgColorReg") Then cmb_Disp2EmgColorReg(2).text = rs!Content
            ElseIf (rs!name = "LANE3_Disp1EmgColorGuest") Then cmb_Disp1EmgColorGuest(2).text = rs!Content
            ElseIf (rs!name = "LANE3_Disp2EmgColorGuest") Then cmb_Disp2EmgColorGuest(2).text = rs!Content
            ElseIf (rs!name = "LANE3_Disp1EmgColorNoRec") Then cmb_Disp1EmgColorNoRec(2).text = rs!Content
            ElseIf (rs!name = "LANE3_Disp2EmgColorNoRec") Then cmb_Disp2EmgColorNoRec(2).text = rs!Content
            ElseIf (rs!name = "LANE3_Disp1EmgColorBKList") Then cmb_Disp1EmgColorBKList(2).text = rs!Content
            ElseIf (rs!name = "LANE3_Disp2EmgColorBKList") Then cmb_Disp2EmgColorBKList(2).text = rs!Content
            ElseIf (rs!name = "LANE3_Disp1EmgColorTaxi") Then cmb_Disp1EmgColorTaxi(2).text = rs!Content
            ElseIf (rs!name = "LANE3_Disp2EmgColorTaxi") Then cmb_Disp2EmgColorTaxi(2).text = rs!Content
            ElseIf (rs!name = "LANE3_Disp1EmgColorDay") Then cmb_Disp1EmgColorDay(2).text = rs!Content
            ElseIf (rs!name = "LANE3_Disp2EmgColorDay") Then cmb_Disp2EmgColorDay(2).text = rs!Content
            ElseIf (rs!name = "LANE3_Disp1EmgColorRegExpDate") Then cmb_Disp1EmgColorRegExpDate(2).text = rs!Content
            ElseIf (rs!name = "LANE3_Disp2EmgColorRegExpDate") Then cmb_Disp2EmgColorRegExpDate(2).text = rs!Content
            ElseIf (rs!name = "LANE3_Disp1EmgColorGuestRegCar") Then cmb_Disp1EmgColorGuestRegCar(2).text = rs!Content
            ElseIf (rs!name = "LANE3_Disp2EmgColorGuestRegCar") Then cmb_Disp2EmgColorGuestRegCar(2).text = rs!Content
            ElseIf (rs!name = "LANE3_Disp1EmgColorGuestRegCarExpDate") Then cmb_Disp1EmgColorGuestRegCarExpDate(2).text = rs!Content
            ElseIf (rs!name = "LANE3_Disp2EmgColorGuestRegCarExpDate") Then cmb_Disp2EmgColorGuestRegCarExpDate(2).text = rs!Content
            
            ElseIf (rs!name = "LANE4_Disp1EmgColorReg") Then cmb_Disp1EmgColorReg(3).text = rs!Content
            ElseIf (rs!name = "LANE4_Disp2EmgColorReg") Then cmb_Disp2EmgColorReg(3).text = rs!Content
            ElseIf (rs!name = "LANE4_Disp1EmgColorGuest") Then cmb_Disp1EmgColorGuest(3).text = rs!Content
            ElseIf (rs!name = "LANE4_Disp2EmgColorGuest") Then cmb_Disp2EmgColorGuest(3).text = rs!Content
            ElseIf (rs!name = "LANE4_Disp1EmgColorNoRec") Then cmb_Disp1EmgColorNoRec(3).text = rs!Content
            ElseIf (rs!name = "LANE4_Disp2EmgColorNoRec") Then cmb_Disp2EmgColorNoRec(3).text = rs!Content
            ElseIf (rs!name = "LANE4_Disp1EmgColorBKList") Then cmb_Disp1EmgColorBKList(3).text = rs!Content
            ElseIf (rs!name = "LANE4_Disp2EmgColorBKList") Then cmb_Disp2EmgColorBKList(3).text = rs!Content
            ElseIf (rs!name = "LANE4_Disp1EmgColorTaxi") Then cmb_Disp1EmgColorTaxi(3).text = rs!Content
            ElseIf (rs!name = "LANE4_Disp2EmgColorTaxi") Then cmb_Disp2EmgColorTaxi(3).text = rs!Content
            ElseIf (rs!name = "LANE4_Disp1EmgColorDay") Then cmb_Disp1EmgColorDay(3).text = rs!Content
            ElseIf (rs!name = "LANE4_Disp2EmgColorDay") Then cmb_Disp2EmgColorDay(3).text = rs!Content
            ElseIf (rs!name = "LANE4_Disp1EmgColorRegExpDate") Then cmb_Disp1EmgColorRegExpDate(3).text = rs!Content
            ElseIf (rs!name = "LANE4_Disp2EmgColorRegExpDate") Then cmb_Disp2EmgColorRegExpDate(3).text = rs!Content
            ElseIf (rs!name = "LANE4_Disp1EmgColorGuestRegCar") Then cmb_Disp1EmgColorGuestRegCar(3).text = rs!Content
            ElseIf (rs!name = "LANE4_Disp2EmgColorGuestRegCar") Then cmb_Disp2EmgColorGuestRegCar(3).text = rs!Content
            ElseIf (rs!name = "LANE4_Disp1EmgColorGuestRegCarExpDate") Then cmb_Disp1EmgColorGuestRegCarExpDate(3).text = rs!Content
            ElseIf (rs!name = "LANE4_Disp2EmgColorGuestRegCarExpDate") Then cmb_Disp2EmgColorGuestRegCarExpDate(3).text = rs!Content
            
            ElseIf (rs!name = "LANE5_Disp1EmgColorReg") Then cmb_Disp1EmgColorReg(4).text = rs!Content
            ElseIf (rs!name = "LANE5_Disp2EmgColorReg") Then cmb_Disp2EmgColorReg(4).text = rs!Content
            ElseIf (rs!name = "LANE5_Disp1EmgColorGuest") Then cmb_Disp1EmgColorGuest(4).text = rs!Content
            ElseIf (rs!name = "LANE5_Disp2EmgColorGuest") Then cmb_Disp2EmgColorGuest(4).text = rs!Content
            ElseIf (rs!name = "LANE5_Disp1EmgColorNoRec") Then cmb_Disp1EmgColorNoRec(4).text = rs!Content
            ElseIf (rs!name = "LANE5_Disp2EmgColorNoRec") Then cmb_Disp2EmgColorNoRec(4).text = rs!Content
            ElseIf (rs!name = "LANE5_Disp1EmgColorBKList") Then cmb_Disp1EmgColorBKList(4).text = rs!Content
            ElseIf (rs!name = "LANE5_Disp2EmgColorBKList") Then cmb_Disp2EmgColorBKList(4).text = rs!Content
            ElseIf (rs!name = "LANE5_Disp1EmgColorTaxi") Then cmb_Disp1EmgColorTaxi(4).text = rs!Content
            ElseIf (rs!name = "LANE5_Disp2EmgColorTaxi") Then cmb_Disp2EmgColorTaxi(4).text = rs!Content
            ElseIf (rs!name = "LANE5_Disp1EmgColorDay") Then cmb_Disp1EmgColorDay(4).text = rs!Content
            ElseIf (rs!name = "LANE5_Disp2EmgColorDay") Then cmb_Disp2EmgColorDay(4).text = rs!Content
            ElseIf (rs!name = "LANE5_Disp1EmgColorRegExpDate") Then cmb_Disp1EmgColorRegExpDate(4).text = rs!Content
            ElseIf (rs!name = "LANE5_Disp2EmgColorRegExpDate") Then cmb_Disp2EmgColorRegExpDate(4).text = rs!Content
            ElseIf (rs!name = "LANE5_Disp1EmgColorGuestRegCar") Then cmb_Disp1EmgColorGuestRegCar(4).text = rs!Content
            ElseIf (rs!name = "LANE5_Disp2EmgColorGuestRegCar") Then cmb_Disp2EmgColorGuestRegCar(4).text = rs!Content
            ElseIf (rs!name = "LANE5_Disp1EmgColorGuestRegCarExpDate") Then cmb_Disp1EmgColorGuestRegCarExpDate(4).text = rs!Content
            ElseIf (rs!name = "LANE5_Disp2EmgColorGuestRegCarExpDate") Then cmb_Disp2EmgColorGuestRegCarExpDate(4).text = rs!Content
            
            ElseIf (rs!name = "LANE6_Disp1EmgColorReg") Then cmb_Disp1EmgColorReg(5).text = rs!Content
            ElseIf (rs!name = "LANE6_Disp2EmgColorReg") Then cmb_Disp2EmgColorReg(5).text = rs!Content
            ElseIf (rs!name = "LANE6_Disp1EmgColorGuest") Then cmb_Disp1EmgColorGuest(5).text = rs!Content
            ElseIf (rs!name = "LANE6_Disp2EmgColorGuest") Then cmb_Disp2EmgColorGuest(5).text = rs!Content
            ElseIf (rs!name = "LANE6_Disp1EmgColorNoRec") Then cmb_Disp1EmgColorNoRec(5).text = rs!Content
            ElseIf (rs!name = "LANE6_Disp2EmgColorNoRec") Then cmb_Disp2EmgColorNoRec(5).text = rs!Content
            ElseIf (rs!name = "LANE6_Disp1EmgColorBKList") Then cmb_Disp1EmgColorBKList(5).text = rs!Content
            ElseIf (rs!name = "LANE6_Disp2EmgColorBKList") Then cmb_Disp2EmgColorBKList(5).text = rs!Content
            ElseIf (rs!name = "LANE6_Disp1EmgColorTaxi") Then cmb_Disp1EmgColorTaxi(5).text = rs!Content
            ElseIf (rs!name = "LANE6_Disp2EmgColorTaxi") Then cmb_Disp2EmgColorTaxi(5).text = rs!Content
            ElseIf (rs!name = "LANE6_Disp1EmgColorDay") Then cmb_Disp1EmgColorDay(5).text = rs!Content
            ElseIf (rs!name = "LANE6_Disp2EmgColorDay") Then cmb_Disp2EmgColorDay(5).text = rs!Content
            ElseIf (rs!name = "LANE6_Disp1EmgColorRegExpDate") Then cmb_Disp1EmgColorRegExpDate(5).text = rs!Content
            ElseIf (rs!name = "LANE6_Disp2EmgColorRegExpDate") Then cmb_Disp2EmgColorRegExpDate(5).text = rs!Content
            ElseIf (rs!name = "LANE6_Disp1EmgColorGuestRegCar") Then cmb_Disp1EmgColorGuestRegCar(5).text = rs!Content
            ElseIf (rs!name = "LANE6_Disp2EmgColorGuestRegCar") Then cmb_Disp2EmgColorGuestRegCar(5).text = rs!Content
            ElseIf (rs!name = "LANE6_Disp1EmgColorGuestRegCarExpDate") Then cmb_Disp1EmgColorGuestRegCarExpDate(5).text = rs!Content
            ElseIf (rs!name = "LANE6_Disp2EmgColorGuestRegCarExpDate") Then cmb_Disp2EmgColorGuestRegCarExpDate(5).text = rs!Content
            
            End If
            
            rs.MoveNext
        Loop
    
        Set rs = Nothing
    End If
End Sub


Private Sub cmd_Snd_Default_Click()

    Dim i As Integer
    
    For i = 0 To MAX_LANE_COUNT - 1
        txt_Str_Reg(i).text = "등록"
        txt_Str_Guest(i).text = "미등록"
        txt_Str_NoRec(i).text = "미인식"
        txt_Str_BKList(i).text = "출입제한"
        txt_Str_Taxi(i).text = "영업차량"
        txt_Str_Day(i).text = "요일위반"
        txt_Str_RegExpDate(i).text = "기간만료"
        txt_Str_GuestRegCar(i).text = "방문예약"
        txt_Str_GuestRegCarExpDate(i).text = "방문예약만료"
        
        Tmp_SNDFILE_Reg(i) = App.Path & "\sound\Bell.wav"
        Tmp_SNDFILE_Guest(i) = App.Path & "\sound\Bell.wav"
        Tmp_SNDFILE_NoRec(i) = App.Path & "\sound\Bell.wav"
        Tmp_SNDFILE_BlackList(i) = App.Path & "\sound\Bell.wav"
        Tmp_SNDFILE_Taxi(i) = App.Path & "\sound\Bell.wav"
        Tmp_SNDFILE_Day(i) = App.Path & "\sound\Bell.wav"
        Tmp_SNDFILE_RegExpDate(i) = App.Path & "\sound\Bell.wav"
        Tmp_SNDFILE_GuestRegCar(i) = App.Path & "\sound\Bell.wav"
        Tmp_SNDFILE_GuestRegCarExpDate(i) = App.Path & "\sound\Bell.wav"
        
        chk_SND_Reg(i).value = 0
        chk_SND_Guest(i).value = 0
        chk_SND_NoRec(i).value = 0
        chk_SND_BKList(i).value = 1
        chk_SND_Taxi(i).value = 0
        chk_SND_Day(i).value = 0
        chk_SND_RegExpDate(i).value = 0
        chk_SND_GuestRegCar(i).value = 0
        chk_SND_GuestRegCarExpDate(i).value = 0
    Next i
End Sub




Private Sub btn_PLY_BKList_Click(Index As Integer)
    Call Sound_Out(Tmp_SNDFILE_BlackList(Index))
End Sub

Private Sub btn_PLY_Guest_Click(Index As Integer)
    Call Sound_Out(Tmp_SNDFILE_Guest(Index))
End Sub

Private Sub btn_PLY_NoRec_Click(Index As Integer)
    Call Sound_Out(Tmp_SNDFILE_NoRec(Index))
End Sub

Private Sub btn_PLY_Reg_Click(Index As Integer)
    Call Sound_Out(Tmp_SNDFILE_Reg(Index))
End Sub

Private Sub btn_PLY_Taxi_Click(Index As Integer)
    Call Sound_Out(Tmp_SNDFILE_Taxi(Index))
End Sub

Private Sub btn_PLY_Day_Click(Index As Integer)
    Call Sound_Out(Tmp_SNDFILE_Day(Index))
End Sub

Private Sub btn_PLY_RegExpDate_Click(Index As Integer)
    Call Sound_Out(Tmp_SNDFILE_RegExpDate(Index))
End Sub

Private Sub btn_SND_Reg_Click(Index As Integer)
    Dim tmpFileName As String
On Error GoTo Err_p
    
        tmpFileName = Format(Now, "YYYYMMDD_HHMMSS")
        tmpFileName = App.Path & "\Sound\*.wav"
            
        CommonDialog1.CancelError = True
        CommonDialog1.InitDir = App.Path & "\Sound\"
        CommonDialog1.Filter = "녹음파일(*.wav)|*.wav"
        CommonDialog1.fileName = tmpFileName
        CommonDialog1.ShowOpen
        tmpFileName = CommonDialog1.fileName
        tmpFileName = Mid(tmpFileName, 1, Len(tmpFileName) - 4) & ".wav"
        
        Tmp_SNDFILE_Reg(Index) = tmpFileName
        
        'Debug.Print Tmp_SNDFILE_Reg(Index) ' 임시 테스트
    
    Exit Sub

Err_p:
        Select Case Err
            Case 32755 '  Dialog Cancelled
                'MsgBox "you cancelled the dialog box"
            Case Else
                'MsgBox "Unexpected error. Err " & Err & " : " & Error
        End Select
End Sub
Private Sub btn_SND_BKList_Click(Index As Integer)
    
    Dim tmpFileName As String
On Error GoTo Err_p
    
        tmpFileName = Format(Now, "YYYYMMDD_HHMMSS")
        tmpFileName = App.Path & "\Sound\*.wav"
            
        CommonDialog1.CancelError = True
        CommonDialog1.InitDir = App.Path & "\Sound\"
        CommonDialog1.Filter = "녹음파일(*.wav)|*.wav"
        CommonDialog1.fileName = tmpFileName
        CommonDialog1.ShowOpen
        tmpFileName = CommonDialog1.fileName
        tmpFileName = Mid(tmpFileName, 1, Len(tmpFileName) - 4) & ".wav"
        
        Tmp_SNDFILE_BlackList(Index) = tmpFileName
        
        'Debug.Print Tmp_SNDFILE_BlackList(Index) ' 임시 테스트
    
    Exit Sub

Err_p:
        Select Case Err
            Case 32755 '  Dialog Cancelled
                'MsgBox "you cancelled the dialog box"
            Case Else
                'MsgBox "Unexpected error. Err " & Err & " : " & Error
        End Select
End Sub

Private Sub btn_SND_Guest_Click(Index As Integer)
    
    Dim tmpFileName As String
On Error GoTo Err_p
    
        tmpFileName = Format(Now, "YYYYMMDD_HHMMSS")
        tmpFileName = App.Path & "\Sound\*.wav"
            
        CommonDialog1.CancelError = True
        CommonDialog1.InitDir = App.Path & "\Sound\"
        CommonDialog1.Filter = "녹음파일(*.wav)|*.wav"
        CommonDialog1.fileName = tmpFileName
        CommonDialog1.ShowOpen
        tmpFileName = CommonDialog1.fileName
        tmpFileName = Mid(tmpFileName, 1, Len(tmpFileName) - 4) & ".wav"
        
        Tmp_SNDFILE_Guest(Index) = tmpFileName
        
        'Debug.Print Tmp_SNDFILE_Guest(Index) ' 임시 테스트
    
    Exit Sub

Err_p:
        Select Case Err
            Case 32755 '  Dialog Cancelled
                'MsgBox "you cancelled the dialog box"
            Case Else
                'MsgBox "Unexpected error. Err " & Err & " : " & Error
        End Select
End Sub

Private Sub btn_SND_NoRec_Click(Index As Integer)
    
    Dim tmpFileName As String
On Error GoTo Err_p
    
        tmpFileName = Format(Now, "YYYYMMDD_HHMMSS")
        tmpFileName = App.Path & "\Sound\*.wav"
            
        CommonDialog1.CancelError = True
        CommonDialog1.InitDir = App.Path & "\Sound\"
        CommonDialog1.Filter = "녹음파일(*.wav)|*.wav"
        CommonDialog1.fileName = tmpFileName
        CommonDialog1.ShowOpen
        tmpFileName = CommonDialog1.fileName
        tmpFileName = Mid(tmpFileName, 1, Len(tmpFileName) - 4) & ".wav"
        
        Tmp_SNDFILE_NoRec(Index) = tmpFileName
        
        'Debug.Print Tmp_SNDFILE_NoRec(Index) ' 임시 테스트
    
    Exit Sub

Err_p:
        Select Case Err
            Case 32755 '  Dialog Cancelled
                'MsgBox "you cancelled the dialog box"
            Case Else
                'MsgBox "Unexpected error. Err " & Err & " : " & Error
        End Select
End Sub


Private Sub btn_SND_Taxi_Click(Index As Integer)
    
        Dim tmpFileName As String
On Error GoTo Err_p
    
        tmpFileName = Format(Now, "YYYYMMDD_HHMMSS")
        tmpFileName = App.Path & "\Sound\*.wav"
            
        CommonDialog1.CancelError = True
        CommonDialog1.InitDir = App.Path & "\Sound\"
        CommonDialog1.Filter = "녹음파일(*.wav)|*.wav"
        CommonDialog1.fileName = tmpFileName
        CommonDialog1.ShowOpen
        tmpFileName = CommonDialog1.fileName
        tmpFileName = Mid(tmpFileName, 1, Len(tmpFileName) - 4) & ".wav"
        
        Tmp_SNDFILE_Taxi(Index) = tmpFileName
        
        'Debug.Print Tmp_SNDFILE_Taxi(Index) ' 임시 테스트
    
    Exit Sub

Err_p:
        Select Case Err
            Case 32755 '  Dialog Cancelled
                'MsgBox "you cancelled the dialog box"
            Case Else
                'MsgBox "Unexpected error. Err " & Err & " : " & Error
        End Select
End Sub



Private Sub btn_SND_Day_Click(Index As Integer)
        Dim tmpFileName As String
On Error GoTo Err_p
    
        tmpFileName = Format(Now, "YYYYMMDD_HHMMSS")
        tmpFileName = App.Path & "\Sound\*.wav"
            
        CommonDialog1.CancelError = True
        CommonDialog1.InitDir = App.Path & "\Sound\"
        CommonDialog1.Filter = "녹음파일(*.wav)|*.wav"
        CommonDialog1.fileName = tmpFileName
        CommonDialog1.ShowOpen
        tmpFileName = CommonDialog1.fileName
        tmpFileName = Mid(tmpFileName, 1, Len(tmpFileName) - 4) & ".wav"
        
        Tmp_SNDFILE_Day(Index) = tmpFileName
        
        'Debug.Print Tmp_SNDFILE_Day(Index) ' 임시 테스트
    
    Exit Sub

Err_p:
        Select Case Err
            Case 32755 '  Dialog Cancelled
                'MsgBox "you cancelled the dialog box"
            Case Else
                'MsgBox "Unexpected error. Err " & Err & " : " & Error
        End Select
End Sub


Private Sub btn_SND_RegExpDate_Click(Index As Integer)
        Dim tmpFileName As String
On Error GoTo Err_p
    
        tmpFileName = Format(Now, "YYYYMMDD_HHMMSS")
        tmpFileName = App.Path & "\Sound\*.wav"
            
        CommonDialog1.CancelError = True
        CommonDialog1.InitDir = App.Path & "\Sound\"
        CommonDialog1.Filter = "녹음파일(*.wav)|*.wav"
        CommonDialog1.fileName = tmpFileName
        CommonDialog1.ShowOpen
        tmpFileName = CommonDialog1.fileName
        tmpFileName = Mid(tmpFileName, 1, Len(tmpFileName) - 4) & ".wav"
        
        Tmp_SNDFILE_RegExpDate(Index) = tmpFileName
        
        'Debug.Print Tmp_SNDFILE_Day(Index) ' 임시 테스트
    
    Exit Sub

Err_p:
        Select Case Err
            Case 32755 '  Dialog Cancelled
                'MsgBox "you cancelled the dialog box"
            Case Else
                'MsgBox "Unexpected error. Err " & Err & " : " & Error
        End Select
End Sub


Private Sub Load_MainStr_Config()

    Dim i As Integer
    
    For i = 0 To MAX_LANE_COUNT - 1
        
        If (Glo_Display_Direct = "가로") Then
            TXT_MAX_LENGTH = 12
            txt_Str_Reg(i).ToolTipText = "한글 6글자 제한"
            txt_Str_Guest(i).ToolTipText = "한글 6글자 제한"
            txt_Str_NoRec(i).ToolTipText = "한글 6글자 제한"
            txt_Str_BKList(i).ToolTipText = "한글 6글자 제한"
            txt_Str_Taxi(i).ToolTipText = "한글 6글자 제한"
            txt_Str_Day(i).ToolTipText = "한글 6글자 제한"
            txt_Str_RegExpDate(i).ToolTipText = "한글 6글자 제한"
            txt_Str_GuestRegCar(i).ToolTipText = "한글 6글자 제한"
            txt_Str_GuestRegCarExpDate(i).ToolTipText = "한글 6글자 제한"
        
        ElseIf (Glo_Display_Direct = "세로") Then
            TXT_MAX_LENGTH = 24
            txt_Str_Reg(i).ToolTipText = "한글 12글자 제한.두 줄로 나눠서 출력됩니다. 짝수로 입력하세요."
            txt_Str_Guest(i).ToolTipText = "한글 12글자 제한.두 줄로 나눠서 출력됩니다. 짝수로 입력하세요."
            txt_Str_NoRec(i).ToolTipText = "한글 12글자 제한.두 줄로 나눠서 출력됩니다. 짝수로 입력하세요."
            txt_Str_BKList(i).ToolTipText = "한글 12글자 제한.두 줄로 나눠서 출력됩니다. 짝수로 입력하세요."
            txt_Str_Taxi(i).ToolTipText = "한글 12글자 제한.두 줄로 나눠서 출력됩니다. 짝수로 입력하세요."
            txt_Str_Day(i).ToolTipText = "한글 12글자 제한.두 줄로 나눠서 출력됩니다. 짝수로 입력하세요."
            txt_Str_RegExpDate(i).ToolTipText = "한글 12글자 제한.두 줄로 나눠서 출력됩니다. 짝수로 입력하세요."
            txt_Str_GuestRegCar(i).ToolTipText = "한글 12글자 제한.두 줄로 나눠서 출력됩니다. 짝수로 입력하세요."
            txt_Str_GuestRegCarExpDate(i).ToolTipText = "한글 12글자 제한.두 줄로 나눠서 출력됩니다. 짝수로 입력하세요."
        End If
        
        txt_Str_Reg(i).text = Glo_Str_Reg(i)
        txt_Str_Guest(i).text = Glo_Str_Guest(i)
        txt_Str_NoRec(i).text = Glo_Str_NoRec(i)
        txt_Str_BKList(i).text = Glo_Str_BlackList(i)
        txt_Str_Taxi(i).text = Glo_Str_Taxi(i)
        txt_Str_Day(i).text = Glo_Str_Day(i)
        txt_Str_RegExpDate(i).text = Glo_Str_RegExpDate(i)
        txt_Str_GuestRegCar(i).text = Glo_Str_GuestRegCar(i)
        txt_Str_GuestRegCarExpDate(i).text = Glo_Str_GuestRegCarExpDate(i)
    Next i

End Sub


Private Sub Save_DispString_Config()

    Dim i As Integer
    
    For i = 0 To MAX_LANE_COUNT - 1
        Glo_Str_Reg(i) = txt_Str_Reg(i).text
        Glo_Str_Guest(i) = txt_Str_Guest(i).text
        Glo_Str_NoRec(i) = txt_Str_NoRec(i).text
        Glo_Str_BlackList(i) = txt_Str_BKList(i).text
        Glo_Str_Taxi(i) = txt_Str_Taxi(i).text
        Glo_Str_Day(i) = txt_Str_Day(i).text
        Glo_Str_RegExpDate(i) = txt_Str_RegExpDate(i).text
        Glo_Str_GuestRegCar(i) = txt_Str_GuestRegCar(i).text
        Glo_Str_GuestRegCarExpDate(i) = txt_Str_GuestRegCarExpDate(i).text
    Next i

    Call Put_Ini("System Config", "Str_Lane1_Reg", Glo_Str_Reg(0))
    Call Put_Ini("System Config", "Str_Lane1_Guest", Glo_Str_Guest(0))
    Call Put_Ini("System Config", "Str_Lane1_NoRec", Glo_Str_NoRec(0))
    Call Put_Ini("System Config", "Str_Lane1_BlackList", Glo_Str_BlackList(0))
    Call Put_Ini("System Config", "Str_Lane1_Taxi", Glo_Str_Taxi(0))
    Call Put_Ini("System Config", "Str_Lane1_Day", Glo_Str_Day(0))
    Call Put_Ini("System Config", "Str_Lane1_RegExpDate", Glo_Str_RegExpDate(0))
    Call Put_Ini("System Config", "Str_Lane1_GuestRegCar", Glo_Str_GuestRegCar(0))
    Call Put_Ini("System Config", "Str_Lane1_GuestRegCarExpDate", Glo_Str_GuestRegCarExpDate(0))
    
    Call Put_Ini("System Config", "Str_Lane2_Reg", Glo_Str_Reg(1))
    Call Put_Ini("System Config", "Str_Lane2_Guest", Glo_Str_Guest(1))
    Call Put_Ini("System Config", "Str_Lane2_NoRec", Glo_Str_NoRec(1))
    Call Put_Ini("System Config", "Str_Lane2_BlackList", Glo_Str_BlackList(1))
    Call Put_Ini("System Config", "Str_Lane2_Taxi", Glo_Str_Taxi(1))
    Call Put_Ini("System Config", "Str_Lane2_Day", Glo_Str_Day(1))
    Call Put_Ini("System Config", "Str_Lane2_RegExpDate", Glo_Str_RegExpDate(1))
    Call Put_Ini("System Config", "Str_Lane2_GuestRegCar", Glo_Str_GuestRegCar(1))
    Call Put_Ini("System Config", "Str_Lane2_GuestRegCarExpDate", Glo_Str_GuestRegCarExpDate(1))
    Call Put_Ini("System Config", "Str_Lane3_Reg", Glo_Str_Reg(2))
    Call Put_Ini("System Config", "Str_Lane3_Guest", Glo_Str_Guest(2))
    Call Put_Ini("System Config", "Str_Lane3_NoRec", Glo_Str_NoRec(2))
    Call Put_Ini("System Config", "Str_Lane3_BlackList", Glo_Str_BlackList(2))
    Call Put_Ini("System Config", "Str_Lane3_Taxi", Glo_Str_Taxi(2))
    Call Put_Ini("System Config", "Str_Lane3_Day", Glo_Str_Day(2))
    Call Put_Ini("System Config", "Str_Lane3_RegExpDate", Glo_Str_RegExpDate(2))
    Call Put_Ini("System Config", "Str_Lane3_GuestRegCar", Glo_Str_GuestRegCar(2))
    Call Put_Ini("System Config", "Str_Lane3_GuestRegCarExpDate", Glo_Str_GuestRegCarExpDate(2))
    Call Put_Ini("System Config", "Str_Lane4_Reg", Glo_Str_Reg(3))
    Call Put_Ini("System Config", "Str_Lane4_Guest", Glo_Str_Guest(3))
    Call Put_Ini("System Config", "Str_Lane4_NoRec", Glo_Str_NoRec(3))
    Call Put_Ini("System Config", "Str_Lane4_BlackList", Glo_Str_BlackList(3))
    Call Put_Ini("System Config", "Str_Lane4_Taxi", Glo_Str_Taxi(3))
    Call Put_Ini("System Config", "Str_Lane4_Day", Glo_Str_Day(3))
    Call Put_Ini("System Config", "Str_Lane4_RegExpDate", Glo_Str_RegExpDate(3))
    Call Put_Ini("System Config", "Str_Lane4_GuestRegCar", Glo_Str_GuestRegCar(3))
    Call Put_Ini("System Config", "Str_Lane4_GuestRegCarExpDate", Glo_Str_GuestRegCarExpDate(3))
    Call Put_Ini("System Config", "Str_Lane5_Reg", Glo_Str_Reg(4))
    Call Put_Ini("System Config", "Str_Lane5_Guest", Glo_Str_Guest(4))
    Call Put_Ini("System Config", "Str_Lane5_NoRec", Glo_Str_NoRec(4))
    Call Put_Ini("System Config", "Str_Lane5_BlackList", Glo_Str_BlackList(4))
    Call Put_Ini("System Config", "Str_Lane5_Taxi", Glo_Str_Taxi(4))
    Call Put_Ini("System Config", "Str_Lane5_Day", Glo_Str_Day(4))
    Call Put_Ini("System Config", "Str_Lane5_RegExpDate", Glo_Str_RegExpDate(4))
    Call Put_Ini("System Config", "Str_Lane5_GuestRegCar", Glo_Str_GuestRegCar(4))
    Call Put_Ini("System Config", "Str_Lane5_GuestRegCarExpDate", Glo_Str_GuestRegCarExpDate(4))
    Call Put_Ini("System Config", "Str_Lane6_Reg", Glo_Str_Reg(5))
    Call Put_Ini("System Config", "Str_Lane6_Guest", Glo_Str_Guest(5))
    Call Put_Ini("System Config", "Str_Lane6_NoRec", Glo_Str_NoRec(5))
    Call Put_Ini("System Config", "Str_Lane6_BlackList", Glo_Str_BlackList(5))
    Call Put_Ini("System Config", "Str_Lane6_Taxi", Glo_Str_Taxi(5))
    Call Put_Ini("System Config", "Str_Lane6_Day", Glo_Str_Day(5))
    Call Put_Ini("System Config", "Str_Lane6_RegExpDate", Glo_Str_RegExpDate(5))
    Call Put_Ini("System Config", "Str_Lane6_GuestRegCar", Glo_Str_GuestRegCar(5))
    Call Put_Ini("System Config", "Str_Lane6_GuestRegCarExpDate", Glo_Str_GuestRegCarExpDate(5))
End Sub

Private Sub Save_DispColor_Config()

    Dim upColor As Byte
    Dim downColor As Byte
    Dim i, Index As Integer
    Dim Text1Color As String
    Dim Text2Color As String
    Dim KeyValue1, KeyValue2 As String
    
    
    For Index = 0 To MAX_LANE_COUNT - 1
            For i = 0 To 9 - 1 '등록/미등록/미인식/출입제한/영업차량/요일제위반/기간만료/방문예약/방문예약만료
                If (i = 0) Then
                    Text1Color = cmb_Disp1EmgColorReg(Index).text
                    Text2Color = cmb_Disp2EmgColorReg(Index).text
                ElseIf (i = 1) Then
                    Text1Color = cmb_Disp1EmgColorGuest(Index).text
                    Text2Color = cmb_Disp2EmgColorGuest(Index).text
                ElseIf (i = 2) Then
                    Text1Color = cmb_Disp1EmgColorNoRec(Index).text
                    Text2Color = cmb_Disp2EmgColorNoRec(Index).text
                ElseIf (i = 3) Then
                    Text1Color = cmb_Disp1EmgColorBKList(Index).text
                    Text2Color = cmb_Disp2EmgColorBKList(Index).text
                ElseIf (i = 4) Then
                    Text1Color = cmb_Disp1EmgColorTaxi(Index).text
                    Text2Color = cmb_Disp2EmgColorTaxi(Index).text
                ElseIf (i = 5) Then
                    Text1Color = cmb_Disp1EmgColorDay(Index).text
                    Text2Color = cmb_Disp2EmgColorDay(Index).text
                ElseIf (i = 6) Then
                    Text1Color = cmb_Disp1EmgColorRegExpDate(Index).text
                    Text2Color = cmb_Disp2EmgColorRegExpDate(Index).text
                ElseIf (i = 7) Then
                    Text1Color = cmb_Disp1EmgColorGuestRegCar(Index).text
                    Text2Color = cmb_Disp2EmgColorGuestRegCar(Index).text
                ElseIf (i = 8) Then
                    Text1Color = cmb_Disp1EmgColorGuestRegCarExpDate(Index).text
                    Text2Color = cmb_Disp2EmgColorGuestRegCarExpDate(Index).text
                End If
                
                upColor = GetDispColorData(Text1Color)
                downColor = GetDispColorData(Text2Color)
                
                
                
                Select Case Index
                Case 0
                    KeyValue1 = "LANE1"
                    KeyValue2 = "LANE1"
                Case 1
                    KeyValue1 = "LANE2"
                    KeyValue2 = "LANE2"
                Case 2
                    KeyValue1 = "LANE3"
                    KeyValue2 = "LANE3"
                Case 3
                    KeyValue1 = "LANE4"
                    KeyValue2 = "LANE4"
                Case 4
                    KeyValue1 = "LANE5"
                    KeyValue2 = "LANE5"
                Case 5
                    KeyValue1 = "LANE6"
                    KeyValue2 = "LANE6"
                End Select
                
                Select Case i
                    Case 0
                        KeyValue1 = KeyValue1 & "_Disp1EmgColorReg"
                        KeyValue2 = KeyValue2 & "_Disp2EmgColorReg"
                        Glo_Disp1_Reg(Index) = upColor
                        Glo_Disp2_Reg(Index) = downColor
                    Case 1
                        KeyValue1 = KeyValue1 & "_Disp1EmgColorGuest"
                        KeyValue2 = KeyValue2 & "_Disp2EmgColorGuest"
                        Glo_Disp1_Guest(Index) = upColor
                        Glo_Disp2_Guest(Index) = downColor
                    Case 2
                        KeyValue1 = KeyValue1 & "_Disp1EmgColorNoRec"
                        KeyValue2 = KeyValue2 & "_Disp2EmgColorNoRec"
                        Glo_Disp1_NoRec(Index) = upColor
                        Glo_Disp2_NoRec(Index) = downColor
                    Case 3
                        KeyValue1 = KeyValue1 & "_Disp1EmgColorBKList"
                        KeyValue2 = KeyValue2 & "_Disp2EmgColorBKList"
                        Glo_Disp1_BlackList(Index) = upColor
                        Glo_Disp2_BlackList(Index) = downColor
                    Case 4
                        KeyValue1 = KeyValue1 & "_Disp1EmgColorTaxi"
                        KeyValue2 = KeyValue2 & "_Disp2EmgColorTaxi"
                        Glo_Disp1_Taxi(Index) = upColor
                        Glo_Disp2_Taxi(Index) = downColor
                    Case 5
                        KeyValue1 = KeyValue1 & "_Disp1EmgColorDay"
                        KeyValue2 = KeyValue2 & "_Disp2EmgColorDay"
                        Glo_Disp1_Day(Index) = upColor
                        Glo_Disp2_Day(Index) = downColor
                    Case 6
                        KeyValue1 = KeyValue1 & "_Disp1EmgColorRegExpDate"
                        KeyValue2 = KeyValue2 & "_Disp2EmgColorRegExpDate"
                        Glo_Disp1_RegExpDate(Index) = upColor
                        Glo_Disp2_RegExpDate(Index) = downColor
                    Case 7
                        KeyValue1 = KeyValue1 & "_Disp1EmgColorGuestRegCar"
                        KeyValue2 = KeyValue2 & "_Disp2EmgColorGuestRegCar"
                        Glo_Disp1_GuestRegCar(Index) = upColor
                        Glo_Disp2_GuestRegCar(Index) = downColor
                    Case 8
                        KeyValue1 = KeyValue1 & "_Disp1EmgColorGuestRegCarExpDate"
                        KeyValue2 = KeyValue2 & "_Disp1EmgColorGuestRegCarExpDate"
                        Glo_Disp1_GuestRegCarExpDate(Index) = upColor
                        Glo_Disp2_GuestRegCarExpDate(Index) = downColor
                        
                End Select
                        
                adoConn.Execute "UPDATE tb_config set Content = '" & Text1Color & "' WHERE NAME = '" & KeyValue1 & "'"
                adoConn.Execute "UPDATE tb_config set Content = '" & Text2Color & "' WHERE NAME = '" & KeyValue2 & "'"

                
            Next i
    Next Index
    
End Sub


Private Sub Command1_Click()
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '사운드 설정
    Call Save_Sound_Config
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Call Save_DispString_Config
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Call Save_DispColor_Config
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub


Private Sub Form_Load()

    Left = (Screen.width - width) / 2   ' 폼을 가로로 중앙에 놓습니다.
    Top = (Screen.height - height) / 2   ' 폼을 세로로 중앙에 놓습니다.
    
    '사운드 설정
    Call Load_Sound_Config
    
    Call Load_MainStr_Config
    
    Call Load_Disp_Config
    
    
End Sub


Private Sub chk_SOUND_YN_Click()
    Dim i As Integer
    
    If (chk_SOUND_YN.value = 1) Then
        'frmSOUND.Enabled = True
        For i = 0 To MAX_LANE_COUNT - 1
            chk_SND_Reg(i).Enabled = True
            chk_SND_Guest(i).Enabled = True
            chk_SND_NoRec(i).Enabled = True
            chk_SND_BKList(i).Enabled = True
            chk_SND_Taxi(i).Enabled = True
            chk_SND_Day(i).Enabled = True
            chk_SND_RegExpDate(i).Enabled = True
            chk_SND_GuestRegCar(i).Enabled = True
            chk_SND_GuestRegCarExpDate(i).Enabled = True
        Next i
    Else
        'frmSOUND.Enabled = False
        For i = 0 To MAX_LANE_COUNT - 1
            chk_SND_Reg(i).Enabled = False
            chk_SND_Guest(i).Enabled = False
            chk_SND_NoRec(i).Enabled = False
            chk_SND_BKList(i).Enabled = False
            chk_SND_Taxi(i).Enabled = False
            chk_SND_Day(i).Enabled = False
            chk_SND_RegExpDate(i).Enabled = False
            chk_SND_GuestRegCar(i).Enabled = False
            chk_SND_GuestRegCarExpDate(i).Enabled = False
            
        Next i
    End If
    
    
    
    
    
    
End Sub


