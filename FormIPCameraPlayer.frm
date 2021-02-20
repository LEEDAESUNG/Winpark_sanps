VERSION 5.00
Object = "{DF2BBE39-40A8-433B-A279-073F48DA94B6}#1.0#0"; "axvlc.dll"
Begin VB.Form FormIPCameraPlayer 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ParkingManager™11145"
   ClientHeight    =   9450
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12150
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9450
   ScaleWidth      =   12150
   StartUpPosition =   3  'Windows 기본값
   Begin VB.PictureBox Picture1 
      Height          =   540
      Left            =   11220
      ScaleHeight     =   480
      ScaleWidth      =   495
      TabIndex        =   2
      Top             =   3975
      Width           =   555
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1080
      Left            =   11085
      TabIndex        =   1
      Top             =   2340
      Width           =   1065
   End
   Begin AXVLCCtl.VLCPlugin2 VLCPlugin21 
      Height          =   9450
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10995
      AutoLoop        =   0   'False
      AutoPlay        =   -1  'True
      Toolbar         =   -1  'True
      ExtentWidth     =   19394
      ExtentHeight    =   16669
      MRL             =   ""
      Object.Visible         =   -1  'True
      Volume          =   100
      StartTime       =   0
      BaseURL         =   ""
      BackColor       =   16777215
      FullscreenEnabled=   -1  'True
      Branding        =   -1  'True
   End
   Begin VB.Line Line1 
      X1              =   6570
      X2              =   10860
      Y1              =   0
      Y2              =   0
   End
End
Attribute VB_Name = "FormIPCameraPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Vopt(1) As String

'이미지 캡쳐 저장
Private Sub Command1_Click()
    Dim resp As Variant
    VLCPlugin21.playlist.pause
    Set resp = VLCPlugin21.video.takeSnapshot
    Picture1.Picture = resp
    VLCPlugin21.playlist.Play
End Sub

Private Sub Form_Load()
    Dim resp As Variant
    VLCPlugin21.playlist.Stop
    VLCPlugin21.playlist.Items.Clear
    VLCPlugin21.playlist.Play
End Sub

Public Sub Play(url As String, name As String)
    'Vopt(0) = ":sout=#transcode{vcodec=mp4v,vb=1024,scale=1,acodec=mp3,ab=128,channels=2}:duplicate{dst=std{access=file,mux=ps,url=c:\" & name & ".mpg}}"
    'Vopt(0) = ":--sout #duplicate{dst=display,dst=std{access=file,mux=ps,dst=c:\" & name & ".mpg}}"
    'Vopt(0) = ":deinterlace-mode=x"
    'VLCPlugin21.audio.mute = True
    'VLCPlugin21.playlist.Add url, "TEST", Vopt(0)
    
    Debug.Print "VLC version:" & VLCPlugin21.getVersionInfo
    
    VLCPlugin21.playlist.Add url
    VLCPlugin21.playlist.Play
    
    Me.Caption = Me.Caption & " - " & name
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    VLCPlugin21.playlist.Stop
    VLCPlugin21.playlist.Items.Clear
    'Unload Me
End Sub

Private Sub Form_Resize()
    VLCPlugin21.Top = 0
    VLCPlugin21.Left = 25
    VLCPlugin21.width = Me.width - 170
    VLCPlugin21.height = Me.height + 50
End Sub

