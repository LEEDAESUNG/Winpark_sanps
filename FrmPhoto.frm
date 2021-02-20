VERSION 5.00
Begin VB.Form FrmPhoto 
   Caption         =   "ParkingManager¢â"
   ClientHeight    =   5475
   ClientLeft      =   10890
   ClientTop       =   5385
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   8880
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Ãë ¼Ò"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   7365
      TabIndex        =   4
      Top             =   4680
      Width           =   1320
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "È® ÀÎ"
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   12
         Charset         =   129
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   5910
      TabIndex        =   3
      Top             =   4680
      Width           =   1320
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3210
      Left            =   2610
      TabIndex        =   2
      Top             =   1065
      Width           =   2940
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3240
      Left            =   225
      TabIndex        =   1
      Top             =   1050
      Width           =   2340
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "±¼¸²"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   210
      TabIndex        =   0
      Top             =   675
      Width           =   2370
   End
   Begin VB.Label lbl_PhotoPath 
      BeginProperty Font 
         Name            =   "³ª´®°íµñ"
         Size            =   9.75
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   225
      TabIndex        =   5
      Top             =   195
      Width           =   8445
   End
   Begin VB.Image Image1 
      Height          =   3000
      Left            =   6015
      Picture         =   "FrmPhoto.frx":0000
      Stretch         =   -1  'True
      Top             =   1155
      Width           =   2505
   End
End
Attribute VB_Name = "FrmPhoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If (Right(lbl_PhotoPath.Caption, 4) = ".jpg" Or Right(lbl_PhotoPath.Caption, 4) = ".JPG") Then
        Frm_Canon.txt_PhotoPath.Text = lbl_PhotoPath.Caption
        Unload Me
    End If
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
    If Right(File1.Path, 1) = "\" Then
        lbl_PhotoPath.Caption = File1.Path & File1.filename
    Else
        lbl_PhotoPath.Caption = File1.Path & "\" & File1.filename
    End If
    
    If (Right(lbl_PhotoPath.Caption, 4) = ".jpg" Or Right(lbl_PhotoPath.Caption, 4) = ".JPG") Then
        Image1.Picture = LoadPicture(lbl_PhotoPath.Caption)
    End If

End Sub

Private Sub cmd_Button_Click(Index As Integer)
        Call Err_doc(Format(Now, "yyyy-mm-dd hh:nn:ss") & "    RFID ÀÏ°ýµî·Ï Á¾·á")
        Unload Me
        Exit Sub
End Sub

Private Sub Form_Load()

    lbl_PhotoPath.Caption = ""



End Sub
