VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents MyText As TextBox
Attribute MyText.VB_VarHelpID = -1

Private Sub MyText_GotFocus()
    MyText.SelStart = 0
    MyText.SelLength = Len(MyText.Text)
    
End Sub

Private Sub MyText_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        SendKeys "{TAB}", True
        KeyCode = 0
    ElseIf KeyCode = vbKeyUp Then
        SendKeys "+{TAB}", True
        KeyCode = 0
    End If
End Sub

Private Sub MyText_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}", True
        KeyAscii = 0
    End If
End Sub
