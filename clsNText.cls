VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNtext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents NText As TextBox
Attribute NText.VB_VarHelpID = -1

Private Sub NText_GotFocus()
    NText.Text = Val(NText.Text)
    NText.SelStart = 0
    NText.SelLength = Len(NText.Text)
    
End Sub

Private Sub NText_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        SendKeys "{TAB}", True
        KeyCode = 0
    ElseIf KeyCode = vbKeyUp Then
        SendKeys "+{TAB}", True
        KeyCode = 0
    End If
End Sub

Private Sub NText_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}", True
        KeyAscii = 0
    End If
End Sub

Private Sub NText_Change()
If (NText.Text = "") Then
Else
    If Not ((IsNumeric(Right(NText.Text, 1)) Or (Right(NText.Text, 1) = "."))) Then
            NText.Text = Left(NText.Text, Len(NText.Text) - 1)
            NText.SelStart = Len(NText.Text)
    End If
End If
End Sub

Private Sub NText_LostFocus()
NText.Text = Val(NText.Text)
End Sub
