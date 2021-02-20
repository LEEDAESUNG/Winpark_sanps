Attribute VB_Name = "Module1"
Option Explicit


Public Const LVM_FIRST = &H1000
Public Const LVM_SETCOLUMNWIDTH = (LVM_FIRST + 30)
Public Const LVSCW_AUTOSIZE = -1
Public Const LVSCW_AUTOSIZE_USEHEADER = -2

Public Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55
Public Const Lvm_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54
Public Const LVS_EX_FULLROWSELECT = &H20

Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long



Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Const INFINITE = -1&
Public Const SYNCHRONIZE = &H100000









Public Sub ListViewExtended(ListView As ListView)
    Dim dwExStyle As Long
    Dim nRT As Long

    dwExStyle = SendMessage(ListView.hwnd, LVM_GETEXTENDEDLISTVIEWSTYLE, 0, 0)
    nRT = SendMessage(ListView.hwnd, Lvm_SETEXTENDEDLISTVIEWSTYLE, 0, dwExStyle Or LVS_EX_FULLROWSELECT)

End Sub

Public Sub makeexcel(lv As ListView, fileName As String, Header As String)
    Dim ecdata As New ExcelFile
    Dim i As Long
    Dim j As Long

'리스트뷰를 엑셀 데이터로 변환 저장한다.
With ecdata

    '파일명 설정 및 생성
    'filename$ = ".\" & filename & ".xls"
    .CreateFile fileName$
    '.setfilepassword "암호"
    .PrintGridLines = False
    '여백 설정
    .SetMargin xlsTopMargin, 1.5
    .SetMargin xlsLeftMargin, 1.5
    .SetMargin xlsRightMargin, 1.5
    .SetMargin xlsBottomMargin, 1.5
    '기본 높이 설정
    '.setdefaultrowheight 14
    .SetFont "arial", 10, xlsNoFormat              'font0
    .SetFont "arial", 10, xlsBold                  'font1
    .SetFont "arial", 10, xlsBold                  'font2
    .SetFont "impact", 15, xlsBold

    .SetColumnWidth 1, 5, 20
    '.setrowheight 1, 17
    '.setrowheight 2, 30
    
    .SetHeader "kiba"
    .SetFooter "kiba"
    .WriteValue xlsText, xlsFont3, xlsCentreAlign, xlsNormal, 1, 1, Trim(Header)
    For i = 1 To lv.ColumnHeaders.Count

        .WriteValue xlsText, xlsFont2, xlsCentreAlign, xlsNormal, 3, i, lv.ColumnHeaders.Item(i).text

    Next i
    For i = 1 To lv.ListItems.Count
        .SetColumnWidth 1, lv.ColumnHeaders.Count, 20
        For j = 1 To lv.ColumnHeaders.Count
            If j = 1 Then
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, i + 3, j, lv.ListItems(i).text
            Else
                .WriteValue xlsText, xlsFont0, xlsCentreAlign, xlsNormal, i + 3, j, lv.ListItems(i).SubItems(j - 1)
            End If
        Next j
    Next i
    .CloseFile
End With
MsgBox "저장이 완료되었습니다."


End Sub


'Windows 10 Resolved Permission deny
Public Sub Sendkeys(text$, Optional wait As Boolean = False)
    Dim WshShell As Object
    Set WshShell = CreateObject("wscript.shell")
    WshShell.Sendkeys text, wait
    Set WshShell = Nothing
End Sub


Public Function StringToUTF8BytesArray(ByVal sText As String) As Byte()
    On Error GoTo Err
    
    Dim adoStream As ADODB.Stream
    Dim convertData() As Byte
    
    Set adoStream = New ADODB.Stream
    adoStream.mode = adModeReadWrite
    adoStream.Type = adTypeText
    adoStream.Charset = "utf-8"
    adoStream.Open
    adoStream.WriteText sText
    adoStream.Flush
    
    adoStream.Position = 0
    adoStream.Type = adTypeBinary
    adoStream.Read 3
    convertData = adoStream.Read()
    
    adoStream.Close
    StringToUTF8BytesArray = convertData
    
    Exit Function
    
Err:
    Call DataLogger("[StringToUTF8ByteArray] Err:" & Err.Description)
End Function

Public Function UTF8ByteArrayToString(ByRef byteArray() As Byte) As String
    On Error GoTo Err
    
    Dim adoStream As ADODB.Stream
    Dim sConvertText As String
    
    Set adoStream = New ADODB.Stream
    adoStream.mode = adModeReadWrite
    adoStream.Type = adTypeText
    adoStream.Charset = "utf-8"
    adoStream.Open
    adoStream.WriteText byteArray
    adoStream.Flush
    
    adoStream.Position = 0
    adoStream.Type = adTypeText
    sConvertText = adoStream.ReadText
    
    adoStream.Close
    UTF8ByteArrayToString = sConvertText
    Exit Function
Err:
    Call DataLogger("[UTF8ByteArrayToString] Err:" & Err.Description)
End Function

