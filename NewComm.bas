Attribute VB_Name = "Module5"
Option Explicit

Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Declare Function SetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Public Sub Sound_Out(s As String)
'Dim Ret As Long
'Dim S_file As String
'
'S_file = Dir(App.Path & "\Sound\" & s)
'If (S_file <> "") Then
'    Ret = PlaySound(App.Path & "\Sound\" & s, 0, &H1 Or &H40)
'End If
Dim Ret As Long
Dim S_file As String

S_file = Dir(s)
If (S_file <> "") Then
    Ret = PlaySound(s, 0, &H1 Or &H40)
End If
End Sub

Public Sub Err_doc(Err_str As String)
Dim intFileNum As Integer
intFileNum = FreeFile()
Open Doc_Path_Name$ & Format(Now, "yyyy-mm-dd") & ".doc" For Append As #intFileNum
Print #intFileNum, Format(Now, "yyyy-mm-dd hh:nn:ss ") & Err_str
Close #intFileNum
End Sub

Public Function Get_Ini(App_Name As String, Key_Name As String, Default_Value As String) As String
Dim RetVal As String * 200
Dim tmp
tmp = GetPrivateProfileString(App_Name, Key_Name, Default_Value, RetVal, LenH(RetVal), IniFileName$)
If (tmp = 0) Then
    Get_Ini = Default_Value
Else
    Get_Ini = RTrim(LeftH$(RetVal, tmp))
End If
End Function
Public Sub Put_Ini(App_Name As String, Key_Name As String, Put_Data As String)
Dim tmp
tmp = WritePrivateProfileString(App_Name, Key_Name, Put_Data, IniFileName$)
End Sub
Public Function Get_Ini2(App_Name As String, Key_Name As String, Default_Value As String, InitFileName As String) As String
Dim RetVal As String * 200
Dim tmp
tmp = GetPrivateProfileString(App_Name, Key_Name, Default_Value, RetVal, LenH(RetVal), InitFileName)
If (tmp = 0) Then
    Get_Ini2 = Default_Value
Else
    Get_Ini2 = RTrim(LeftH$(RetVal, tmp))
End If
End Function
Public Sub Put_Ini2(App_Name As String, Key_Name As String, Put_Data As String, InitFileName As String)
Dim tmp
tmp = WritePrivateProfileString(App_Name, Key_Name, Put_Data, InitFileName)
End Sub

Public Function NoneSTX_Bcc_Chk(Indata As String) As Byte
' 이함수는 순수 DATA 만 CHECK SUM한 결과를 리턴한다.
Dim i As Integer
Dim Cnt As Integer
Dim Hex_Str As String
Dim Bcc_Val
Cnt = Len(Indata)
Hex_Str = ""
For i = 1 To Cnt
    If (Len(Hex(Asc(Mid(Indata, i, 1)))) = 1) Then
        Hex_Str = Hex_Str & "0" & Hex(Asc(Mid(Indata, i, 1)))
    Else
        Hex_Str = Hex_Str & Hex(Asc(Mid(Indata, i, 1)))
    End If
Next i
Cnt = Len(Hex_Str) / 2
Cnt = Len(Hex_Str) / 2
Bcc_Val = 0
For i = 1 To Cnt
    Bcc_Val = Bcc_Val Xor Val("&H" & Mid(Hex_Str, (i * 2) - 1, 2))
Next i
Bcc_Val = Bcc_Val Xor 3
NoneSTX_Bcc_Chk = Bcc_Val
End Function

Public Sub None_Delay_Time(DTime As Single)
Dim PauseTime As Single
Dim start  As Single
PauseTime = DTime
start = Timer
Do While Timer < start + PauseTime
    If (Timer < start) Then
        start = start - 86400
    End If
Loop
End Sub

Public Sub Delay_Time(DTime As Single)
Dim PauseTime As Single
Dim start  As Single
PauseTime = DTime
start = Timer
Do While Timer < start + PauseTime
    DoEvents
    If (Timer < start) Then
        start = start - 86400
    End If
Loop
End Sub

' 설  명: 문자열(s)의 chk_pos번째 문자의 한글/영문 판단
' 복귀값: 0 = 영문, 1 = 한글 첫번째 바이트, 2 = 한글 두번째 바이트
Private Function WhatByte(ByVal s As String, ByVal chk_pos As Integer) As Integer
Dim i As Integer
  '******************** 에러 처리 *********************
  If chk_pos > LenH(s) Then WhatByte = 0: Exit Function
  s = StrConv(s, 128)  '한글 코드 페이지
  For i = 1 To chk_pos
     If AscB(MidB(s, i, 1)) >= 128 Then
       WhatByte = 1: i = i + 1
     Else
       WhatByte = 0
     End If
  Next i

  If WhatByte = 1 And (i - 1) = chk_pos Then WhatByte = 2

End Function

' 설  명: 문자열(s)의 길이를 구한다.
' 복귀값: 한글은 2바이트, 영문은 1바이트로 계산하여 전체 문자열의 길이를 구한다.
Public Function LenH(ByVal s As String) As Integer
  LenH = LenB(StrConv(s, 128))
End Function

' 설  명: 문자열(s)의 왼쪽부터 n바이트 길이만큼 뽑아낸다.
Public Function LeftH(ByVal s As String, ByVal n As Integer) As String
Dim i, flag As Integer
  '***************** 에러 처리 *****************
  If s = "" Or n <= 0 Then Exit Function
  If n >= LenH(s) Then LeftH = s: Exit Function
  If WhatByte(s, n) = 1 Then n = n - 1: flag = 1
  s = StrConv(s, 128) '한글 코드 페이지.
  For i = 1 To n
     LeftH = LeftH & ChrB(AscB(MidB(s, i, 1)))
  Next i
  If flag Then LeftH = LeftH & ChrB(32)
  LeftH = StrConv(LeftH, 64) '유니 코드로 바꾼다.
End Function

' 설  명: 문자열(s)의 start번째부터 n바이트 길이만큼 뽑아낸다.
Public Function MidH(ByVal s As String, ByVal start As Integer, ByVal n As Integer) As String
Dim flag, fin, i As Integer
  '******************** 에러 처리 ********************
  If s = "" Or start <= 0 Or n <= 0 Then Exit Function
  fin = start + n - 1
  If fin >= LenH(s) Then fin = LenH(s)
  If WhatByte(s, start) = 2 Then
    MidH = ChrB(32): start = start + 1
  End If
  If WhatByte(s, fin) = 1 Then fin = fin - 1: flag = 1
  s = StrConv(s, 128) '한글 코드 페이지.
  For i = start To fin
     MidH = MidH & ChrB(AscB(MidB(s, i, 1)))
  Next i
  If flag Then MidH = MidH & ChrB(32)
  MidH = StrConv(MidH, 64) '유니 코드로 바꾼다.
End Function

' 설  명: 문자열(s)의 오른쪽부터 n바이트 길이만큼 뽑아낸다.
Public Function RightH(ByVal s As String, ByVal n As Integer)
Dim start, fin, i As Integer
  '***************** 에러 처리 *****************
  If s = "" Or n <= 0 Then Exit Function
  If n >= LenH(s) Then RightH = s: Exit Function
  fin = LenH(s)
  start = fin - n + 1
  If WhatByte(s, start) = 2 Then
    RightH = ChrB(32): start = start + 1
  End If
  s = StrConv(s, 128) '한글 코드 페이지.
  For i = start To fin
     RightH = RightH & ChrB(AscB(MidB(s, i, 1)))
  Next i
  RightH = StrConv(RightH, 64) '유니 코드로 바꾼다.
End Function

' 설  명: 문자열(s)의 중간에 있는 공백을 몽땅 없앤다.
Public Function XTrim(ByVal s As String) As String
Dim i As Integer
Dim Length As Integer
  Length = Len(s)
  For i = 1 To Length
     If Mid(s, i, 1) <> " " Then
       XTrim = XTrim & Mid(s, i, 1)
     End If
  Next i
End Function

' 설  명: 문자열(X  ===> ?)문자로 치환한다...
Public Function XToQ(ByVal s As String) As String
Dim i As Integer
Dim Length As Integer
  Length = Len(s)
  For i = 1 To Length
     If Mid(s, i, 1) <> "X" Then
       XToQ = XToQ & Mid(s, i, 1)
     Else
        XToQ = XToQ & "?"
     End If
  Next i
End Function

' 설  명: 문자열(X  ===> ?)문자로 치환한다...
Public Function QToAll(ByVal s As String) As String
Dim i As Integer
Dim Length As Integer
  Length = Len(s)
  For i = 1 To Length
     If Mid(s, i, 1) <> "?" Then
       QToAll = QToAll & Mid(s, i, 1)
     Else
        QToAll = QToAll & "*"
     End If
  Next i
End Function

Public Function IsChar(ByVal s As String) As Integer
Dim i As Integer
Dim num As Integer
Dim Length As Integer
  Length = Len(s)
  For i = 1 To Length
     If Mid(s, i, 1) = "X" Then
       num = num + 1
     End If
  Next i
  IsChar = num
End Function


Public Function ByteArrayToHex(ByRef ByteArray() As Byte) As String
    Dim lb As Long, ub As Long
    Dim l As Long, strRet As String
    Dim lonRetLen As Long, lonPos As Long
    Dim strHex As String, lonLenHex As Long
    
    lb = LBound(ByteArray)
    ub = UBound(ByteArray)
    lonRetLen = ((ub - lb) + 1) * 3
    strRet = Space$(lonRetLen)
    lonPos = 1
    
    For l = lb To ub
        strHex = Hex$(ByteArray(l))
        If Len(strHex) = 1 Then strHex = "0" & strHex
        If l <> ub Then
            Mid$(strRet, lonPos, 3) = strHex & " "
            lonPos = lonPos + 3
        Else
            Mid$(strRet, lonPos, 3) = strHex
        End If
    Next l
    
    ByteArrayToHex = strRet
End Function

