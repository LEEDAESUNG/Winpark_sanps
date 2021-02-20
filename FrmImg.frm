VERSION 5.00
Begin VB.Form FrmImg 
   BorderStyle     =   0  '없음
   Caption         =   "ParkingManager™"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9900
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   9900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '소유자 가운데
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFF00&
      Height          =   7170
      Left            =   0
      Top             =   15
      Width           =   9885
   End
   Begin VB.Image Image1 
      Height          =   7170
      Left            =   0
      Stretch         =   -1  'True
      Top             =   15
      Width           =   9885
   End
End
Attribute VB_Name = "FrmImg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
