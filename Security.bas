Attribute VB_Name = "Security"
Option Explicit

Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, ByVal cbData As Long) As Long
Public Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_DYN_DATA = &H80000006
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_USERS = &H80000003
Public Const ERROR_SUCCESS = 0&
Public Const REG_OPTION_NON_VOLATILE = &O0
Public Const KEY_ALL_CLASSES As Long = &HF0063
Public Const KEY_ALL_ACCESS = &H3F
Public Const REG_SZ As Long = 1
Public Const REG_DWORD = 4

Public Const VisPath = "SOFTWARE\JAWOOTEK\Parking"

'Type SYSTEMTIME
'        wYear As Integer
'        wMonth As Integer
'        wDayOfWeek As Integer
'        wDay As Integer
'        wHour As Integer
'        wMinute As Integer
'        wSecond As Integer
'        wMilliseconds As Integer
'End Type
'Declare Function SetLocalTime Lib "KERNEL32" (lpSystemTime As SYSTEMTIME) As Long

Public Sub Time_Sync()
    
    Dim Qry As String
    Dim rs As ADODB.Recordset
    Dim bQryResult As Boolean

On Error GoTo Err_P

    Set rs = New ADODB.Recordset
    Qry = "Select date_format(now()," & Chr(34) & "%Y%m%d%H%i%S" & Chr(34) & ");"
    'rs.Open Qry, adoConn
    bQryResult = DataBaseQuery(rs, adoConn, Qry, False)
    If (bQryResult = False) Then
        FrmTcpServer.ListData.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    네트워크 및 DB 점검바랍니다", 0
        Call DataLogger("[TimeSync]    " & "네트워크 및 DB 점검바랍니다")
        Exit Sub
    End If
    
    If Not (rs.EOF) Then
        'Debug.Print rs(0)
        If (rs(0) <> Format(Now, "yyyymmddhhnn")) Then
            Call Set_Time(rs(0))
            Call DataLogger("DB Time Sync. Success..!!")
        End If
    End If
    Set rs = Nothing

Exit Sub

Err_P:
    Call DataLogger("TimeSync Proc Error : " & Err.Description)

End Sub


Public Sub Set_Time(time_str As String)
    
    Dim Sys_Time As SYSTEMTIME
    Dim tmp As Long
    Dim YYYY As Integer
    Dim MM As Integer
    Dim DD As Integer
    Dim HH As Integer
    Dim NN As Integer
    Dim SS As Integer
    
On Error GoTo Err_P
    
    'yyyymmddhhnnss
    YYYY = Mid(time_str, 1, 4)
    MM = Mid(time_str, 5, 2)
    DD = Mid(time_str, 7, 2)
    HH = Mid(time_str, 9, 2)
    NN = Mid(time_str, 11, 2)
    SS = Mid(time_str, 13, 2)
    Sys_Time.wYear = YYYY
    Sys_Time.wMonth = MM
    Sys_Time.wDayOfWeek = 0
    Sys_Time.wDay = DD
    Sys_Time.wHour = HH
    Sys_Time.wMinute = NN
    Sys_Time.wSecond = SS
    Sys_Time.wMilliseconds = 0
    tmp = SetLocalTime(Sys_Time)

Exit Sub

Err_P:
    Call DataLogger("Set_Time Error")
End Sub


Public Function SetValueEx(ByVal hKey As Long, sValueName As String, lType As Long, vValue As Variant) As Long
    SetValueEx = RegSetValueExString(hKey, sValueName, 0&, lType, vValue, Len(vValue))
End Function

Public Function GetRegValue(TopKey As Long, SubKey As String, ValueTitle As String) As String
    Dim lRetVal As Long
    Dim Buffer As String * 128 '버퍼
    Dim lBufferSize As Long '버퍼크기
    Dim lSubKey As Long
    Dim dType As Long
    Dim i As Integer
    Dim xx As String
    Dim tt As Double

    lBufferSize = 64

    lRetVal = RegOpenKeyEx(TopKey, SubKey, 0, KEY_ALL_ACCESS, lSubKey)
    If lRetVal <> ERROR_SUCCESS Then
        GetRegValue = ""
        RegCloseKey lSubKey
        Exit Function
    End If

   lRetVal = RegQueryValueEx(lSubKey, ValueTitle, 0, dType, ByVal Buffer, lBufferSize)
   If lRetVal = ERROR_SUCCESS Then
        If dType = 4 Then '레지스트리의 값 타입이 Double Word 형이면
            For i = 1 To 4
                xx = Mid(Buffer, i, 1)
                tt = tt + Asc(xx) * 256 ^ (i - 1)
            Next i
                GetRegValue = Trim(tt)
        Else
            xx = ""
            For i = 1 To Len(Buffer)
                If Mid(Buffer, i, 1) = Chr(0) Then Exit For
                xx = xx + Mid(Buffer, i, 1)
            Next i
            GetRegValue = xx
        End If
    Else
        GetRegValue = "" '(원본)레지스트리 키 이름이 없으면 에러가 난다
        'GetRegValue = 0 '(내가고친것)에러나는 이유는? 모르겠음
        RegCloseKey lSubKey
        Exit Function
    End If
    RegCloseKey lSubKey '열려진 레지스트리 키를 닫는다
End Function

Public Function SetRegValue(TopKey As Long, SubKey As String, ValueTitle As String, value As String, dType As Long) As String
    Dim lSubKey As Long
    Dim lRetVal As Long
    Dim iValue As Long

    lRetVal = RegCreateKey(TopKey, SubKey, lSubKey)

    If lRetVal <> ERROR_SUCCESS Then
        SetRegValue = ""
        RegCloseKey lSubKey
        Exit Function
    End If

    If dType = 4 Then '레지스트리의 값 타입이 Double Word 형이면
        iValue = Val(value)
        lRetVal = RegSetValueEx(lSubKey, ValueTitle, 0, REG_DWORD, iValue, 4)
        If lRetVal = ERROR_SUCCESS Then
            SetRegValue = value
        Else
            SetRegValue = ""
        End If
    Else
        If value = "" Then value = " "
        'lRetVal = RegSetValueEx(lSubKey, ValueTitle, 0, REG_SZ, ByVal Value, Len(Value) + 1) 리부팅 시간의 마지막 초가 짤려서 길이를 임으로 늘렸다. 문제?
        lRetVal = RegSetValueEx(lSubKey, ValueTitle, 0, REG_SZ, ByVal value, Len(value) + 2)
        If lRetVal = ERROR_SUCCESS Then
            SetRegValue = value
        Else
            SetRegValue = ""
        End If
    End If
End Function

Public Function GetDriveSerialNumber(Optional ByVal DriveLetter As String) As String


    Dim fso As Object, Drv As Object
    Dim DriveSerial As String
    Dim strtemp As String
    
    'Create a FileSystemObject object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    'Assign the current drive letter if not specified
    If DriveLetter <> "" Then
        Set Drv = fso.GetDrive(DriveLetter)
    Else
        Set Drv = fso.GetDrive(fso.GetDriveName(App.Path))
    End If

    With Drv
        If .IsReady Then
            DriveSerial = Abs(.SerialNumber)
        Else    '"Drive Not Ready!"
            DriveSerial = -1
        End If
    End With
    
    'Clean up
    Set Drv = Nothing
    Set fso = Nothing
    strtemp = Hex(DriveSerial)
    GetDriveSerialNumber = Left(strtemp, 4) & Right(strtemp, 4)
    
End Function

'XOR 알고리즘
Public Function Encrypt(ByRef Original As String) As String
    If LenB(Original) = 0 Or Original = Null Then Exit Function
    
    Dim buf() As Byte
    Dim Key As Byte
    Key = 11
    
    buf() = StrConv(Original, vbFromUnicode)
    
    Dim i As Long
    
    For i = 0 To UBound(buf)
        Encrypt = Encrypt & Right$("0" & Hex$(buf(i) Xor Key), 2)
    Next
End Function

Public Function Decrypt(ByRef Crypted As String) As String
    If LenB(Crypted) = 0 Or Crypted = Null Then Exit Function
    
    Dim i As Long
    Dim Key As Byte
    Key = 11
        
    If Crypted = " " Then
        Exit Function
    End If
        
    For i = 1 To Len(Crypted) Step 2
        Decrypt = Decrypt & ChrB$(CByte("&H" & Mid$(Crypted, i, 2)) Xor Key)
    Next
    
    Decrypt = StrConv(Decrypt, vbUnicode)
End Function

Public Sub HomeLogger(LogStr As String)
    Dim intFileNum As Integer
    
    intFileNum = FreeFile()
    Open Doc_Path_Name$ & Format(Now, "yyyy-mm-dd") & ".txt" For Append As #intFileNum
    Print #intFileNum, "HomeLog_" & Format(Now, "yyyy-mm-dd hh:nn:ss ") & "    " & LogStr
    Close #intFileNum
    FrmTcpServer.ListData.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & LogStr, 0

End Sub

Public Sub DataLogger(LogStr As String)
'Public Sub Err_doc(Err_str As String)
    Dim intFileNum As Integer
    
    intFileNum = FreeFile()
    Open Doc_Path_Name$ & Format(Now, "yyyy-mm-dd") & ".txt" For Append As #intFileNum
    Print #intFileNum, Format(Now, "yyyy-mm-dd hh:nn:ss ") & "    " & LogStr
    Close #intFileNum
    FrmTcpServer.ListData.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & LogStr, 0

End Sub

Public Sub DebugLogger(LogStr As String)
'Public Sub Err_doc(Err_str As String)
    Dim intFileNum As Integer
    
    intFileNum = FreeFile()
    Open App.Path & "\Doc\Debug_" & Format(Now, "yyyy-mm-dd") & ".txt" For Append As #intFileNum
    Print #intFileNum, Format(Now, "yyyy-mm-dd hh:nn:ss ") & "    " & LogStr
    Close #intFileNum
    FrmTcpServer.ListData.AddItem Format(Now, "yyyy-mm-dd hh:nn:ss") & "    " & LogStr, 0

End Sub

Public Function EncodeNDE01(ByVal str As String, ByVal Key As String) As String
    On Error GoTo Err
    
    ' ## 임시 변수 선언
    Dim i As Long, j As Long, Count As Long
    Dim DataTable(4) As String, Result As String, Buffer As Long, Match(2) As Integer
    
    ' ## 바이트 배열 변수 초기화
    Dim Data() As Byte: ReDim Data(Len(str) * 2 - 1)
    
    ' ## 문자열 뒤집기
    Dim ReverseStr As String, ReverseKey As String: ReverseStr = StrReverse(str): ReverseKey = StrReverse(Key)
    
    ' ## DataTable 작성
    DataTable(0) = "y.s1[*m!PR#;J8C6Io<n`w:zB$""D>Sq),?N0lGL@_WfT&794^jv%Ftr{3~kEuUQ]p|=i'K(5dcAVb\Z/a2}xhe+-MHOYXg"
    DataTable(1) = "{;aev>3NqHT2^xMDFBP[#]o/9?EK,m<0ZU.iS\bsOL=6R4(I@G*_kWQ""jg~X5$8+'c17)A}r%Vu-!tdlnhwJCf&|:pyzY`"
    DataTable(2) = "3pM-]V(;Lf$|%sAnl<2.8#U@>+QKy\obWq*FtXk'&dhBjx9rGTYRe=:D6}[NZcJ5,^vHO1IazwEm{_u7)g~S""C`P0i/4?!"
    DataTable(3) = "-""+Fj9=]Crh<\2@JWG7yzw6eq5/ml&v3k,oHD#n}~p(?41Za:IV{R_U8;0td)%N.KPMb!`LOY*f|T'AguX>^$x[cSiEsBQ"
    DataTable(4) = "*_7L~Z?'8F}!>P=-xc[Xs^l$pwG]&4C\h)Yidv1,Wbr`0Nn;RzoT.EHDM:@{k65AJO|m#uV/<f+Baq(Q%yKtjI9S3""e2Ug"
    
    ' ## 뒤집은 평문을 바이트 배열에 대입
    For i = 1 To Len(ReverseStr)
        Buffer = CLng("&H" & Hex$(AscW(Mid$(ReverseStr, i, 1))))
        
        ' ## ASCII / UNICODE 판별
        If Buffer <= &HFF Then
            Data(Count + 1) = Buffer
        Else
            Data(Count) = CLng("&H" & Left$(Hex$(Buffer), 2))
            Data(Count + 1) = CLng("&H" & Right$(Hex$(Buffer), 2))
        End If
        
        Count = Count + 2
    Next i
    
    ' ## 카운트 변수 초기화
    Count = 1
    
    ' ## 키와 바이트 배열을 XOR 연산
    For i = 0 To UBound(Data)
        If Count > Len(ReverseKey) Then
            Data(i) = Data(i) Xor Asc(Mid$(ReverseKey, 1, 1)): Count = 2
        Else
            Data(i) = Data(i) Xor Asc(Mid$(ReverseKey, Count, 1)): Count = Count + 1
        End If
    Next i
    
    ' ## 키의 각 값과 바이트 배열을 XOR 연산
    For i = 0 To UBound(Data)
        For j = 1 To Len(Key)
            Data(i) = Data(i) Xor Asc(Mid$(Key, j, 1))
        Next j
    Next i
    
    ' ## DataTable 과 서로 매칭
    For i = 0 To UBound(Data)
        Match(0) = Data(i) Mod 5: Match(1) = Data(i) \ 94: Match(2) = Data(i) Mod 94
        Result = Result & Mid$(DataTable(Match(0)), Match(1) + 1, 1) & Mid$(DataTable(Match(0)), Match(2) + 1, 1)
    Next i
    
    EncodeNDE01 = StrReverse(Result)
    Exit Function
    
Err:
    EncodeNDE01 = vbNullString
End Function

Public Function DecodeNDE01(ByVal str As String, ByVal Key As String) As String
    On Error GoTo Err
    
    ' ## 임시 변수 선언
    Dim i As Long, j As Long, Count As Long
    Dim DataTable(4) As String, Result As String, Buffer As String, Match As Long
    
    ' ## 바이트 배열 변수 초기화
    Dim Data() As Byte: ReDim Data(Len(str) / 2 - 1)
    
    
    ' ## 문자열 뒤집기
    Dim ReverseStr As String, ReverseKey As String: ReverseStr = StrReverse(str): ReverseKey = StrReverse(Key)
    
    ' ## DataTable 작성
    DataTable(0) = "y.s1[*m!PR#;J8C6Io<n`w:zB$""D>Sq),?N0lGL@_WfT&794^jv%Ftr{3~kEuUQ]p|=i'K(5dcAVb\Z/a2}xhe+-MHOYXg"
    DataTable(1) = "{;aev>3NqHT2^xMDFBP[#]o/9?EK,m<0ZU.iS\bsOL=6R4(I@G*_kWQ""jg~X5$8+'c17)A}r%Vu-!tdlnhwJCf&|:pyzY`"
    DataTable(2) = "3pM-]V(;Lf$|%sAnl<2.8#U@>+QKy\obWq*FtXk'&dhBjx9rGTYRe=:D6}[NZcJ5,^vHO1IazwEm{_u7)g~S""C`P0i/4?!"
    DataTable(3) = "-""+Fj9=]Crh<\2@JWG7yzw6eq5/ml&v3k,oHD#n}~p(?41Za:IV{R_U8;0td)%N.KPMb!`LOY*f|T'AguX>^$x[cSiEsBQ"
    DataTable(4) = "*_7L~Z?'8F}!>P=-xc[Xs^l$pwG]&4C\h)Yidv1,Wbr`0Nn;RzoT.EHDM:@{k65AJO|m#uV/<f+Baq(Q%yKtjI9S3""e2Ug"
    
    ' ## 바이트 배열 복구
    For i = 1 To Len(ReverseStr) Step 2
        Buffer = Mid$(ReverseStr, i, 2)
        
        ' ## DataTable 과 서로 매칭
        For j = 0 To UBound(DataTable)
            If InStr(1, Left$(DataTable(j), 3), Left$(Buffer, 1)) Then Match = j: Exit For
        Next j
        
        ' ## DataTable 에서 원본 값 추출
        Data(Count) = (InStr(1, DataTable(Match), Left$(Buffer, 1)) - 1) * 94 + InStr(1, DataTable(Match), Right$(Buffer, 1)) - 1: Count = Count + 1
    Next i
    
    ' ## 키의 각 값과 바이트 배열을 XOR 연산
    For i = 0 To UBound(Data)
        For j = 1 To Len(Key)
            Data(i) = Data(i) Xor Asc(Mid$(Key, j, 1))
        Next j
    Next i
    
    ' ## 카운트 변수 초기화
    Count = 1
    
    ' ## 키와 바이트 배열을 XOR 연산
    For i = 0 To UBound(Data)
        If Count > Len(ReverseKey) Then
            Data(i) = Data(i) Xor Asc(Mid$(ReverseKey, 1, 1)): Count = 2
        Else
            Data(i) = Data(i) Xor Asc(Mid$(ReverseKey, Count, 1)): Count = Count + 1
        End If
    Next i
    
    ' ## 평문으로 복호화
    For i = 0 To UBound(Data) Step 2
        Buffer = Hex$(Data(i + 1)): If LenB(Buffer) = 2 Then Buffer = "0" & Buffer
        Result = Result & ChrW$(CLng("&H" & Hex$(Data(i)) & Buffer))
    Next i
    
    DecodeNDE01 = StrReverse(Result)
    Exit Function
    
Err:
    DecodeNDE01 = vbNullString
End Function




'C drive 시리얼키 구하기
Private Sub GetClienKey(sKey As String)
    Dim List, Obj, msg
    Dim object
    Dim sStrDrive As String
    
    On Error Resume Next
    
    Set List = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_LogicalDisk") 'HDD고유값
    For Each object In List
        sStrDrive = object.Path_.RelPath
        If (InStr(sStrDrive, "C:") > 0) Then
            'msg = msg & object.VolumeSerialNumber & vbCrLf
            msg = object.VolumeSerialNumber
            Exit For
        End If
    Next
    
    If (Len(msg) > 0) Then
        sKey = msg
    Else
        sKey = "ERROR:C드라이브 오류 입니다."
    End If
    
End Sub
















