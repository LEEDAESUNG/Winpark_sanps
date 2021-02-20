Attribute VB_Name = "POS_Printer"
Option Explicit

Public Declare Sub Out Lib "WIN95IO.DLL" Alias "vbOut" (ByVal nPort As Integer, ByVal nData As Integer)
Public Declare Sub Outw Lib "WIN95IO.DLL" Alias "vbOutw" (ByVal nPort As Integer, ByVal nData As Integer)
Public Declare Function Inp Lib "WIN95IO.DLL" Alias "vbInp" (ByVal nPort As Integer) As Integer
Public Declare Function Inpw Lib "WIN95IO.DLL" Alias "vbInpw" (ByVal nPort As Integer) As Integer

'=================================================================================================================
'WINXP , Win2000
Public Declare Function DlPortReadPortUchar Lib "dlportio.dll" (ByVal Port As Long) As Byte
Public Declare Function DlPortReadPortUshort Lib "dlportio.dll" (ByVal Port As Long) As Integer
Public Declare Function DlPortReadPortUlong Lib "dlportio.dll" (ByVal Port As Long) As Long

Public Declare Sub DlPortReadPortBufferUchar Lib "dlportio.dll" (ByVal Port As Long, Buffer As Any, ByVal Count As Long)
Public Declare Sub DlPortReadPortBufferUshort Lib "dlportio.dll" (ByVal Port As Long, Buffer As Any, ByVal Count As Long)
Public Declare Sub DlPortReadPortBufferUlong Lib "dlportio.dll" (ByVal Port As Long, Buffer As Any, ByVal Count As Long)

Public Declare Sub DlPortWritePortUchar Lib "dlportio.dll" (ByVal Port As Long, ByVal value As Byte)
Public Declare Sub DlPortWritePortUshort Lib "dlportio.dll" (ByVal Port As Long, ByVal value As Integer)
Public Declare Sub DlPortWritePortUlong Lib "dlportio.dll" (ByVal Port As Long, ByVal value As Long)

Public Declare Sub DlPortWritePortBufferUchar Lib "dlportio.dll" (ByVal Port As Long, Buffer As Any, ByVal Count As Long)
Public Declare Sub DlPortWritePortBufferUshort Lib "dlportio.dll" (ByVal Port As Long, Buffer As Any, ByVal Count As Long)
Public Declare Sub DlPortWritePortBufferUlong Lib "dlportio.dll" (ByVal Port As Long, Buffer As Any, ByVal Count As Long)

'3   프린터 에러상태
'4   프린터 준비상태
'5   종이공급 상태
'6   프린터에 문자 도착상태
'7   프린터 작동 상태


Public Declare Function CreateFileNS Lib "Kernel32.dll" Alias "CreateFileA" (ByVal lpFileName As String, _
        ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, _
        ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, _
        ByVal hTemplateFile As Long) As Long
        
'Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, _
'        ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As _
'        SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, _
'        ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long

Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Public Type OVERLAPPED
        Internal As Long
        InternalHigh As Long
        offset As Long
        OffsetHigh As Long
        hEvent As Long
End Type

Public Declare Function LockFile Lib "kernel32" (ByVal hFile As Long, _
        ByVal dwFileOffsetLow As Long, ByVal dwFileOffsetHigh As Long, _
        ByVal nNumberOfBytesToLockLow As Long, ByVal nNumberOfBytesToLockHigh As Long) As Long
Public Declare Function UnlockFile Lib "kernel32" (ByVal hFile As Long, _
        ByVal dwFileOffsetLow As Long, ByVal dwFileOffsetHigh As Long, _
        ByVal nNumberOfBytesToUnlockLow As Long, ByVal nNumberOfBytesToUnlockHigh As Long) As Long

Public Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, _
        ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, _
        ByVal dwMoveMethod As Long) As Long
        
Public Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, _
        lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, _
        lpNumberOfBytesWritten As Long, lpOverlapped As OVERLAPPED) As Long

'// < 버퍼를 문자열로 선언할 경우 >
Public Declare Function WriteFileString Lib "Kernel32.dll" Alias "WriteFile" _
       (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToWrite As Long, _
        lpNumberOfBytesWritten As Long, lpOverlapped As OVERLAPPED) As Long

'// < 화일이 오버랩 속성으로 생성되지 않은 형태로 선언할 경우 >
Public Declare Function WriteFileNO Lib "Kernel32.dll" Alias "WriteFile" _
       (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, _
        lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long

'// < 화일이 오버랩 속성으로 생성되지 않고 버퍼를 문자열로 선언할 경우 >
Public Declare Function WriteFileNOString Lib "Kernel32.dll" Alias "WriteFile" _
       (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToWrite As Long, _
        lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long

Public Declare Function ReadFile Lib "kernel32" _
        (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, _
         lpNumberOfBytesRead As Long, lpOverlapped As OVERLAPPED) As Long

'// < 버퍼를 문자열로 선언할 경우 >
Public Declare Function ReadFileString Lib "Kernel32.dll" Alias "ReadFile" _
       (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, _
        lpNumberOfBytesRead As Long, lpOverlapped As OVERLAPPED) As Long

'// < 화일이 오버랩 속성으로 생성되지 않은 형태로 선언할 경우 >
Public Declare Function ReadFileNO Lib "Kernel32.dll" Alias "ReadFile" _
       (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, _
        lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long

'// < 화일이 오버랩 속성으로 생성되지 않고 버퍼를 문자열로 선언할 경우 >
Public Declare Function ReadFileNOString Lib "Kernel32.dll" Alias "ReadFile" _
       (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, _
        lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long

Public Declare Function RemoveDirectory Lib "kernel32" Alias "RemoveDirectoryA" _
        (ByVal lpPathName As String) As Long

Public Declare Function CloseHandle Lib "kernel32" _
        (ByVal hObject As Long) As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" _
       (ByVal lpFileName As String) As Long
        
Public Declare Function SearchPath Lib "kernel32" Alias "SearchPathA" (ByVal lpPath As String, _
        ByVal lpFileName As String, ByVal lpExtension As String, ByVal nBufferLength As Long, _
        ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
        
Public Declare Function GetFileSize Lib "kernel32" _
        (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Public Declare Function GetFileType Lib "kernel32" _
        (ByVal hFile As Long) As Long

Public Const FILE_TYPE_CHAR = &H2         '/ 문자화일,
Public Const FILE_TYPE_DISK = &H1
Public Const FILE_TYPE_PIPE = &H3
Public Const FILE_TYPE_REMOTE = &H8000
Public Const FILE_TYPE_UNKNOWN = &H0
        
Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Public Declare Function GetFileTime Lib "kernel32" _
        (ByVal hFile As Long, lpCreationTime As FILETIME, _
         lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Public Declare Function FileTimeToLocalFileTime Lib "kernel32" _
        (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Public Declare Function LocalFileTimeToFileTime Lib "kernel32" _
        (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Public Declare Function FileTimeToDosDateTime Lib "kernel32" _
        (lpFileTime As FILETIME, ByVal lpFatDate As Long, ByVal lpFatTime As Long) As Long
Public Declare Function FileTimeToSystemTime Lib "kernel32" _
        (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Public Declare Function SystemTimeToFileTime Lib "kernel32" _
        (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Public Declare Function CompareFileTime Lib "kernel32" _
        (lpFileTime1 As FILETIME, lpFileTime2 As FILETIME) As Long
Public Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, _
        lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, _
        lpLastWriteTime As FILETIME) As Long
       
 Public Const GENERIC_READ = &H80000000
 Public Const GENERIC_WRITE = &H40000000

 Public Const FILE_SHARE_READ = &H1
 Public Const FILE_SHARE_WRITE = &H2

 Public Const CREATE_NEW = 1
 Public Const CREATE_ALWAYS = 2
 Public Const OPEN_ALWAYS = 4
 Public Const OPEN_EXISTING = 3
 Public Const TRUNCATE_EXISTING = 5
 
'/ dwMoveMethod 값
 Public Const FILE_BEGIN = 0     '/ 화일의 시작점
 Public Const FILE_CURRENT = 1   '/ 화일포인터가 있는 현재점
 Public Const FILE_END = 2       '/ EOF(End Of File)

 Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
 Public Const FILE_ATTRIBUTE_COMPRESSED = &H800
 Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
 Public Const FILE_ATTRIBUTE_HIDDEN = &H2
 Public Const FILE_ATTRIBUTE_NORMAL = &H80
 Public Const FILE_ATTRIBUTE_READONLY = &H1
 Public Const FILE_ATTRIBUTE_SYSTEM = &H4
 Public Const FILE_ATTRIBUTE_TEMPORARY = &H100

 Public Const FILE_FLAG_BACKUP_SEMANTICS = &H2000000
 Public Const FILE_FLAG_DELETE_ON_CLOSE = &H4000000
 Public Const FILE_FLAG_NO_BUFFERING = &H20000000
 Public Const FILE_FLAG_OVERLAPPED = &H40000000
 Public Const FILE_FLAG_POSIX_SEMANTICS = &H1000000
 Public Const FILE_FLAG_RANDOM_ACCESS = &H10000000
 Public Const FILE_FLAG_SEQUENTIAL_SCAN = &H8000000
 Public Const FILE_FLAG_WRITE_THROUGH = &H80000000

 Public Const SECURITY_ANONYMOUS_LOGON_RID = &H7
 Public Const SECURITY_CONTEXT_TRACKING = &H40000
 Public Const SECURITY_EFFECTIVE_ONLY = &H80000

 Public Const OFS_MAXPATHNAME = 128
 Public Const OF_CANCEL = &H800
 Public Const OF_CREATE = &H1000
 Public Const OF_DELETE = &H200
 Public Const OF_EXIST = &H4000
 Public Const OF_PARSE = &H100
 Public Const OF_PROMPT = &H2000
 Public Const OF_READ = &H0
 Public Const OF_READWRITE = &H2
 Public Const OF_REOPEN = &H8000
 Public Const OF_SHARE_COMPAT = &H0
 Public Const OF_SHARE_DENY_NONE = &H40
 Public Const OF_SHARE_DENY_READ = &H30
 Public Const OF_SHARE_DENY_WRITE = &H20
 Public Const OF_SHARE_EXCLUSIVE = &H10
 Public Const OF_VERIFY = &H400
 Public Const OF_WRITE = &H1

 '영수증 프린터 타입
 Public Const NO_PRINTER = 0
 Public Const CBM_720 = 1
 Public Const CP_300 = 2
 Public Const TM_300 = 3
 Public Const STAR_300 = 4
 Public Const TM_T88 = 5

Public Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, _
             lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Public Type OFSTRUCT
        cBytes As Byte
        fFixedDisk As Byte
        nErrCode As Integer
        Reserved1 As Integer
        Reserved2 As Integer
        szPathName(OFS_MAXPATHNAME) As Byte
End Type
Public Declare Function QueryPerformanceFrequency Lib "kernel32" _
        (lpFrequency As LARGE_INTEGER) As Long
Public Declare Function QueryPerformanceCounter Lib "kernel32" _
       (lpPerformanceCount As LARGE_INTEGER) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
       
Public Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type


'Public Function Open_Printer(File_Handle As Long) As Boolean
Public Function Open_Printer(File_Handle As Long, Gate As Integer) As Boolean

    Dim Port As String

    If (InStr(1, Glo_Guest_Print_Port(Gate), "COM") > 0) Then
        Port = "Rs232c"
    Else
        Port = Glo_Guest_Print_Port(Gate)
    End If

    Select Case Port
           Case "Rs232c"
                    'If (rs!사용여부) Then
                        If (FrmTcpServer.MSComm(Gate).PortOpen = True) Then
                            Open_Printer = True
                        Else
                            Open_Printer = False
                            Msg_Box.Caption = "Parking System"
                            Msg_Box.Label1.Caption = "영수증 프린터가 Rs232로 설정되어 있으나 포트를 정상적으로 사용할 수 없습니다." & Chr$(13) & Chr$(10) & "환경설정에서 Comm포트설정을 확인하세요!"
                            Msg_Box.Show 1
                        End If
                    'Else
                    '    Open_Printer = False
                    'End If
           Case "LPT1"
                    If (Printer_Status = False) Then
                        Msg_Box.Caption = "Parking System"
                        Msg_Box.Label1.Caption = "프린터 오류 :  영수증 프린터에 이상이 있습니다. 점검 하시기 바랍니다."
                        Msg_Box.Show 1
                        Open_Printer = False
                        Exit Function
                    End If
                    File_Handle = CreateFileNS("Lpt1", GENERIC_WRITE, FILE_SHARE_WRITE, 0, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, 0)
                    If File_Handle = -1 Then
                        Msg_Box.Caption = "Parking System"
                        Msg_Box.Label1.Caption = "영수증 프린터가 LPT1으로 설정되어 있으나  LPT1이 오픈되지 않습니다." & Chr$(13) & Chr$(10) & "시스템을 재부팅 하십시요!"
                        Msg_Box.Show 1
                        Open_Printer = False
                    Else
                        Open_Printer = True
                    End If
           Case "LPT2"
                    If (Printer_Status = False) Then
                        Msg_Box.Caption = "Parking System"
                        Msg_Box.Label1.Caption = "프린터 오류 :  영수증 프린터에 이상이 있습니다. 점검 하시기 바랍니다."
                        Msg_Box.Show 1
                        Open_Printer = False
                        Exit Function
                    End If
                    File_Handle = CreateFileNS("Lpt2", GENERIC_WRITE, FILE_SHARE_WRITE, 0, CREATE_NEW, FILE_ATTRIBUTE_NORMAL, 0)
                    If File_Handle = -1 Then
                        Msg_Box.Caption = "Parking System"
                        Msg_Box.Label1.Caption = "영수증 프린터가 LPT1으로 설정되어 있으나  LPT1이 오픈되지 않습니다." & Chr$(13) & Chr$(10) & "시스템을 재부팅 하십시요!"
                        Msg_Box.Show 1
                        Open_Printer = False
                    Else
                        Open_Printer = True
                    End If
            Case "FILE"
                    Open_Printer = True
                    
            Case "NONE"
                    Open_Printer = False
End Select

End Function

Public Function Print_Type() As Byte

'    Select Case mch
'           Case "CBM-720"
'                tmp = CBM_720
'           Case "CP-300"
'                tmp = CP_300
'           Case "TM-300"
'                tmp = TM_300
'           Case "STAR-300"
'                tmp = STAR_300
'           Case "TM-T88", "WRP-100P"
'                tmp = TM_T88
'           Case Else
'                tmp = NO_PRINTER
'    End Select
'
'    Print_Type = tmp
'
'    Set rs = Nothing

    Print_Type = TM_T88
    
End Function

Public Sub Cash_Draw_Open()
'Dim Rtn As Long
'Dim R As Boolean
'Dim str_Buff As String
'Dim tmp As Byte
'Dim rs As Recordset
''
''Set rs = ParkDb.OpenRecordset("SELECT * FROM 영수증프린터", dbOpenSnapshot)
''If rs.EOF Then
''   Set rs = Nothing
''   Exit Sub
''End If
'
'R = Open_Printer(F_Handle)
'If (R = False) Then
'    Exit Sub
'End If
'
'
'tmp = Print_Type
'Select Case tmp
'       Case CBM_720
'                 str_Buff = Chr$(&H1B) & Chr$(Asc("O"))
'       Case CP_300
'                 str_Buff = Chr$(&H7)
'       Case TM_300
'                 str_Buff = Chr$(&H1B) & Chr$(Asc("p"))
'       Case STAR_300
'                 str_Buff = Chr$(&H7) '또는 str_Buff = Chr$(&H1C)
'       Case TM_T88
'                 str_Buff = Chr$(&H1B) & Chr$(&H70) & Chr$(&H0) & Chr$(&H80) & Chr$(&H80)
'       Case Else
'End Select
'Select Case rs!Port
'           Case "Rs232c"
'                        If (main.MSComm(FEE_PRINT_PORT).PortOpen = True) Then
'                            main.MSComm(FEE_PRINT_PORT).Output = str_Buff
'                        End If
'           Case "LPT1"
'                    Rtn = WriteFileNOString(F_Handle, str_Buff, Len(str_Buff), Len(str_Buff), 0)
'           Case "LPT2"
'                    Rtn = WriteFileNOString(F_Handle, str_Buff, Len(str_Buff), Len(str_Buff), 0)
'End Select
'R = CloseHandle(F_Handle)
End Sub

Public Sub Paper_Cut(Cut_Mode As Byte, Gate As Integer, F_Handle As Long)
'Cut_Mode   1: Partial_Cut  0: Full_Cut
    Dim Rtn As Long
    Dim str_Buff As String
    Dim tmp As Byte
    Dim rs As Recordset
    
    Dim Port As String
    
    tmp = Print_Type
    Select Case tmp
           Case CBM_720
                    str_Buff = Chr$(&H1B) & Chr$(Asc("P")) & Chr$(Cut_Mode)
           Case CP_300
                    str_Buff = Chr$(&H1B) & Chr$(Asc("P")) & Chr$(Asc("0")) '또는 Chr$(&H1B) & Chr$(Asc("R")) & Chr$(Asc("0"))
           Case TM_300
                    str_Buff = Chr$(&H1B) & Chr$(Asc("i"))
           Case STAR_300
                    str_Buff = Chr$(&H1B) & Chr$(Asc("P")) & Chr$(Asc("0")) '모르겠다
           Case TM_T88
                     'str_Buff = Chr$(&H1D) & Chr$(Asc("V")) & Chr$(&H0)
           'Case WRP_100P

                If (Glo_Receipt_Paper_Cut = "1") Then
                    str_Buff = Chr$(&H1B) & Chr$(Asc("m")) & Chr$(&H0)
                Else
                    str_Buff = Chr$(&H1B) & Chr$(Asc("i")) & Chr$(&H0)
                End If
                
           Case Else
    End Select
    Call Paper_Feed(5, Gate, F_Handle)


    If (InStr(1, Glo_Guest_Print_Port(Gate), "COM") > 0) Then
        Port = "Rs232c"
    Else
        Port = Glo_Guest_Print_Port(Gate) 'LPT1 or LPT2 or FILE
    End If
    
    'Select Case rs!Port
    Select Case Port
               Case "Rs232c"
                            If (FrmTcpServer.MSComm(Gate).PortOpen = True) Then
                                FrmTcpServer.MSComm(Gate).Output = str_Buff
                            End If
               Case "LPT1"
                        Rtn = WriteFileNOString(F_Handle, str_Buff, Len(str_Buff), Len(str_Buff), 0)
               Case "LPT2"
                        Rtn = WriteFileNOString(F_Handle, str_Buff, Len(str_Buff), Len(str_Buff), 0)
                Case "FILE"
                         'Call Print_doc("=====================")
                         Call Print_doc("~~~~~~~~~~~~~~~~~~~~~")
    End Select
End Sub

Public Sub Paper_Feed(Feed_Line As Byte, Gate As Integer, F_Handle As Long)
    Dim Rtn As Long
    Dim str_Buff As String
    Dim i As Byte
    Dim tmp As Byte
    Dim rs As Recordset
    Dim Port As String
    
    tmp = Print_Type
    Select Case tmp
           Case CBM_720
                    str_Buff = Chr$(10)
           Case CP_300
                    str_Buff = Chr$(10)
           Case TM_300
                    str_Buff = Chr$(10)
           Case STAR_300
                    str_Buff = Chr$(10)
           Case TM_T88
                     str_Buff = Chr$(10)
           Case Else
    End Select
    
    
    If (InStr(1, Glo_Guest_Print_Port(Gate), "COM")) Then
        Port = "Rs232c"
    Else
        Port = Glo_Guest_Print_Port(Gate)
    End If
    
    Select Case Port
               Case "Rs232c"
                            If (FrmTcpServer.MSComm(Gate).PortOpen = True) Then
                                For i = 1 To Feed_Line
                                    FrmTcpServer.MSComm(Gate).Output = str_Buff
                                Next i
                            End If
               Case "LPT1"
                            For i = 1 To Feed_Line
                                Rtn = WriteFileNOString(F_Handle, str_Buff, Len(str_Buff), Len(str_Buff), 0)
                            Next i
               Case "LPT2"
                            For i = 1 To Feed_Line
                                Rtn = WriteFileNOString(F_Handle, str_Buff, Len(str_Buff), Len(str_Buff), 0)
                            Next i
                Case "FILE"
                            For i = 1 To Feed_Line
                                Call Print_doc("")
                            Next i
    End Select
End Sub

Public Sub Print_String(PrintString As String, mode As Byte, Gate As Integer, F_Handle As Long)
    Dim Rtn As Long
    Dim str_Buff As String
    Dim tmp As Byte
    Dim rs As Recordset
    Dim Port As String
    
    If (InStr(1, Glo_Guest_Print_Port(Gate), "COM") > 0) Then
        Port = "Rs232c"
    Else
        Port = Glo_Guest_Print_Port(Gate)
    End If

tmp = Print_Type
Select Case tmp
       Case CBM_720
                    If (mode = 0) Then
                        str_Buff = Chr$(15)
                    Else
                        str_Buff = Chr$(14)
                    End If
                    Select Case Port
                               Case "Rs232c"
                                            If (FrmTcpServer.MSComm(Gate).PortOpen = True) Then
                                                FrmTcpServer.MSComm(Gate).Output = str_Buff
                                                FrmTcpServer.MSComm(Gate).Output = PrintString & Chr$(10)
                                            End If
                               Case "LPT1"
                                            Rtn = WriteFileNOString(F_Handle, str_Buff, Len(str_Buff), Len(str_Buff), 0)
                                            Rtn = WriteFileNOString(F_Handle, PrintString & Chr$(10), LenH(PrintString & Chr$(10)), LenH(PrintString & Chr$(10)), 0)
                               Case "LPT2"
                                            Rtn = WriteFileNOString(F_Handle, str_Buff, Len(str_Buff), Len(str_Buff), 0)
                                            Rtn = WriteFileNOString(F_Handle, PrintString & Chr$(10), LenH(PrintString & Chr$(10)), LenH(PrintString & Chr$(10)), 0)
                               Case "FILE"
                                            Call Print_doc(PrintString)
                    End Select
       Case CP_300
                    If (mode = 0) Then
                        str_Buff = Chr$(20)
                    Else
                        str_Buff = Chr$(14)
                    End If
                    Select Case Port
                               Case "Rs232c"
                                            If (FrmTcpServer.MSComm(Gate).PortOpen = True) Then
                                                FrmTcpServer.MSComm(Gate).Output = str_Buff
                                                FrmTcpServer.MSComm(Gate).Output = PrintString & Chr$(10)
                                            End If
                               Case "LPT1"
                                            Rtn = WriteFileNOString(F_Handle, str_Buff, Len(str_Buff), Len(str_Buff), 0)
                                            Rtn = WriteFileNOString(F_Handle, PrintString & Chr$(10), LenH(PrintString & Chr$(10)), LenH(PrintString & Chr$(10)), 0)
                               Case "LPT2"
                                            Rtn = WriteFileNOString(F_Handle, str_Buff, Len(str_Buff), Len(str_Buff), 0)
                                            Rtn = WriteFileNOString(F_Handle, PrintString & Chr$(10), LenH(PrintString & Chr$(10)), LenH(PrintString & Chr$(10)), 0)
                               Case "FILE"
                                            Call Print_doc(PrintString)
                    End Select
       Case TM_300
                    If (mode = 0) Then
                        str_Buff = Chr$(&H1B) & Chr$(&H21) & Chr$(&H0)
                    Else
                        str_Buff = Chr$(&H1B) & Chr$(&H21) & Chr$(&H20)
                    End If
                    Select Case Port
                               Case "Rs232c"
                                            If (FrmTcpServer.MSComm(Gate).PortOpen = True) Then
                                                FrmTcpServer.MSComm(Gate).Output = str_Buff
                                                FrmTcpServer.MSComm(Gate).Output = PrintString & Chr$(10)
                                            End If
                               Case "LPT1"
                                            Rtn = WriteFileNOString(F_Handle, str_Buff, Len(str_Buff), Len(str_Buff), 0)
                                            Rtn = WriteFileNOString(F_Handle, PrintString & Chr$(10), LenH(PrintString & Chr$(10)), LenH(PrintString & Chr$(10)), 0)
                               Case "LPT2"
                                            Rtn = WriteFileNOString(F_Handle, str_Buff, Len(str_Buff), Len(str_Buff), 0)
                                            Rtn = WriteFileNOString(F_Handle, PrintString & Chr$(10), LenH(PrintString & Chr$(10)), LenH(PrintString & Chr$(10)), 0)
                               Case "FILE"
                                            Call Print_doc(PrintString)
                    End Select
       Case STAR_300
                    If (mode = 0) Then
                        str_Buff = Chr$(15)
                    Else
                        str_Buff = Chr$(14)
                    End If
                    Select Case Port
                               Case "Rs232c"
                                            If (FrmTcpServer.MSComm(Gate).PortOpen = True) Then
                                                FrmTcpServer.MSComm(Gate).Output = str_Buff
                                                FrmTcpServer.MSComm(Gate).Output = PrintString & Chr$(10)
                                            End If
                               Case "LPT1"
                                            Rtn = WriteFileNOString(F_Handle, str_Buff, Len(str_Buff), Len(str_Buff), 0)
                                            Rtn = WriteFileNOString(F_Handle, PrintString & Chr$(10), LenH(PrintString & Chr$(10)), LenH(PrintString & Chr$(10)), 0)
                               Case "LPT2"
                                            Rtn = WriteFileNOString(F_Handle, str_Buff, Len(str_Buff), Len(str_Buff), 0)
                                            Rtn = WriteFileNOString(F_Handle, PrintString & Chr$(10), LenH(PrintString & Chr$(10)), LenH(PrintString & Chr$(10)), 0)
                               Case "FILE"
                                            Call Print_doc(PrintString)
                    End Select
       Case TM_T88
'''                    If (mode = 0) Then
'''                        str_Buff = Chr$(&H1B) & Chr$(&H21) & Chr$(&H0)
'''                    Else
'''                        str_Buff = Chr$(&H1B) & Chr$(&H21) & Chr$(&H20)
'''                    End If

                    '요금계산기용 영수증프린터
                    If (mode = 0) Then
                        str_Buff = Chr$(&H1D) & Chr$(&H21) & Chr$(&H0)
                    Else
                        str_Buff = Chr$(&H1D) & Chr$(&H21) & Chr$(16)
                    End If
                    Select Case Port
                               Case "Rs232c"
                                            If (FrmTcpServer.MSComm(Gate).PortOpen = True) Then
                                                FrmTcpServer.MSComm(Gate).Output = str_Buff
                                                FrmTcpServer.MSComm(Gate).Output = PrintString & Chr$(10)
                                            End If
                               Case "LPT1"
                                            Rtn = WriteFileNOString(F_Handle, str_Buff, Len(str_Buff), Len(str_Buff), 0)
                                            Rtn = WriteFileNOString(F_Handle, PrintString & Chr$(10), LenH(PrintString & Chr$(10)), LenH(PrintString & Chr$(10)), 0)
                               Case "LPT2"
                                            Rtn = WriteFileNOString(F_Handle, str_Buff, Len(str_Buff), Len(str_Buff), 0)
                                            Rtn = WriteFileNOString(F_Handle, PrintString & Chr$(10), LenH(PrintString & Chr$(10)), LenH(PrintString & Chr$(10)), 0)
                               Case "FILE"
                                            Call Print_doc(PrintString)
                    End Select
       Case Else
End Select

End Sub

Public Function Printer_Status98() As Boolean
   Dim ret1 As Boolean
   Dim ret2 As Boolean
   Dim ret3 As Boolean

   If (Inp(&H379) And &H80) <> &H80 Then
       ret1 = False
   Else
        ret1 = True
   End If


   If ((Inp(&H379) And &H50) <> &H50) Then
       ret2 = False
   Else
        ret2 = True
   End If

   If ((DlPortReadPortUchar(&H379) And &H20) <> &H20) Then
       ret3 = True
   Else
       ret2 = False
   End If

   If (ret1 = True And ret2 = True And ret3 = True) Then
       Printer_Status98 = True
   Else
       Printer_Status98 = False
   End If
End Function
   
Public Function Printer_Status() As Boolean
   Dim ret1 As Boolean
   Dim ret2 As Boolean
   Dim ret3 As Boolean
   
   Printer_Status = True
   Exit Function

   If (DlPortReadPortUchar(&H379) And &H80) <> &H80 Then
       ret1 = False
   Else
        ret1 = True
   End If
   If ((DlPortReadPortUchar(&H379) And &H50) <> &H50) Then
       ret2 = False
   Else
        ret2 = True
   End If

   If ((DlPortReadPortUchar(&H379) And &H20) <> &H20) Then
       ret3 = True
   Else
       ret2 = False
   End If

   If (ret1 = True And ret2 = True And ret3 = True) Then
       Printer_Status = True
   Else
       Printer_Status = False
   End If
End Function
   
Public Sub BarCodePrint(str As String)
'Dim tmp As Boolean
'Dim Title As String
'Dim rs As Recordset
''tmp = Open_Printer(F_Handle)
''If (tmp = False) Then
''    Exit Sub
''End If
'
'Dim ticket_num As String
'Dim Time_Band As String
'
'Set rs = ParkDb.OpenRecordset("SELECT * FROM 인쇄옵션", dbOpenSnapshot)
'
'With main
'
'ticket_num = Get_Ini("시스템", "Reticket_Num", "0001")
'    'Call TDPaper_Cut(1)
'    Call TDPaper_Feed(1)
'    Call TDPrint_String("     [입 차 증]", 1)
'    Call TDPaper_Feed(1)
'
'    Call TDPrint_String("상      호 : " & rs!상호, 0)
'    Call TDPrint_String("주      소 : " & rs!주소, 0)
'    Call TDPrint_String("사업자번호 : " & rs!사업자번호, 0)
'    Call TDPrint_String("근 무 자   : " & WorkManName, 0)
'    Call TDPrint_String("----------------------------------------", 0)
'    Call TDPrint_String("", 0)
'    Call TDPrint_String("입차일자:" & Format(Now, "yyyy-mm-dd"), 1)
'    Call TDPrint_String("입차시간:" & Format(Now, "hh:nn"), 1)
'    Call TDPaper_Feed(1)
'    Call TDPrint_String("티켓번호:" & ticket_num, 1)
'    Call TDPrint_String(Chr$(29) & Chr$(104) & Chr$(100), 0) '바코드 높이 2 ~ 6
'    Call TDPrint_String(Chr$(29) & Chr$(119) & Chr$(3), 0) '바코드 굵기 2 ~ 6
'    Call TDPrint_String(Chr$(29) & "k" & Chr$(70) & Chr$(14) & ticket_num & Format(Now, "yymmddhhnn"), 0) '프린트 바코드
'    Call TDPaper_Feed(1)
'    Call TDPaper_Feed(1)
'    Call TDPaper_Feed(1)
'    Call TDPaper_Feed(1)
'    Call TDPaper_Feed(1)
'
'    Call TDPaper_Cut(1)
''    Time_Band = val(Format(Now, "hh"))
''    ParkDb.Execute "UPDATE 시간대별 SET 일반권입차수 = 일반권입차수 + 1, 데이터유무 = True WHERE 정산여부 = False AND 시간대 = " & Time_Band
''    .DataTime.Refresh
''    .PnlMsg.Caption = "일반권 입차  주차권번호 : " & ticket_num
'    .List1.AddItem "일반권 입차  주차권번호 : " & ticket_num, 0
''    .List1.ListIndex = .List1.ListCount - 1
'    Put_Ini "시스템", "Reticket_Num", Format((val(ticket_num) + 1), "0000")
''    Call Print_String("", 0)
''    Title = "[영 수 증]"
''    Call Print_String(Space((21 - LenH(Title)) / 2) & Title, 1)
''
''    Call Print_String("", 0)
''    Call Print_String("상      호 : " & rs!상호, 0)
''    Call Print_String("주      소 : " & rs!주소, 0)
''    Call Print_String("사업자번호 : " & rs!사업자번호, 0)
''    Call Print_String("근 무 자   : " & WorkManName, 0)
''    Call Print_String("----------------------------------------", 0)
''
''    tmp = CloseHandle(F_Handle)
'
'    Call DataLogger(" 일반권 입차  주차권번호 : " & ticket_num)
'
''    Call None_Delay_Time(Glo_TD_DelayTime)
''    Call Relay_Out(0, 0)
'
'End With

End Sub

Public Sub Print_doc(str As String)
Dim intFileNum As Integer
intFileNum = FreeFile()
Open Db_Path_Name$ & Format(Now, "yyyy-mm-dd") & ".txt" For Append As #intFileNum
Print #intFileNum, str
Close #intFileNum
End Sub

'TD 입차증 발행

Public Sub TDPaper_Cut(Cut_Mode As Byte, F_Handle As Long)
''Cut_Mode   1: Partial_Cut  0: Full_Cut
'Dim Rtn As Long
'Dim str_Buff As String
'Dim tmp As Byte
'Dim rs As Recordset
'
'Set rs = ParkDb.OpenRecordset("SELECT * FROM 영수증프린터", dbOpenSnapshot)
'If rs.EOF Then
'   Set rs = Nothing
'   Exit Sub
'End If
'tmp = Print_Type
'Select Case tmp
'       Case CBM_720
'                str_Buff = Chr$(&H1B) & Chr$(Asc("P")) & Chr$(Cut_Mode)
'       Case CP_300
'                str_Buff = Chr$(&H1B) & Chr$(Asc("P")) & Chr$(Asc("0")) '또는 Chr$(&H1B) & Chr$(Asc("R")) & Chr$(Asc("0"))
'       Case TM_300
'                str_Buff = Chr$(&H1B) & Chr$(Asc("i"))
'       Case STAR_300
'                str_Buff = Chr$(&H1B) & Chr$(Asc("P")) & Chr$(Asc("0")) '모르겠다
'       Case TM_T88
'                 str_Buff = Chr$(&H1D) & Chr$(Asc("V")) & Chr$(&H0)
'       Case Else
'End Select
'Paper_Feed 5
'Select Case rs!Port
'           Case "Rs232c"
'                        If (main.MSComm(TD_PORT).PortOpen = True) Then
'                            main.MSComm(TD_PORT).Output = str_Buff
'                        End If
'           Case "LPT1"
'                    Rtn = WriteFileNOString(F_Handle, str_Buff, Len(str_Buff), Len(str_Buff), 0)
'           Case "LPT2"
'                    Rtn = WriteFileNOString(F_Handle, str_Buff, Len(str_Buff), Len(str_Buff), 0)
'End Select
End Sub

Public Sub TDPaper_Feed(Feed_Line As Byte, F_Handle As Long)
'Dim Rtn As Long
'Dim str_Buff As String
'Dim i As Byte
'Dim tmp As Byte
'Dim rs As Recordset
'
'Set rs = ParkDb.OpenRecordset("SELECT * FROM 영수증프린터", dbOpenSnapshot)
'If rs.EOF Then
'   Set rs = Nothing
'   Exit Sub
'End If
'tmp = Print_Type
'Select Case tmp
'       Case CBM_720
'                str_Buff = Chr$(10)
'       Case CP_300
'                str_Buff = Chr$(10)
'       Case TM_300
'                str_Buff = Chr$(10)
'       Case STAR_300
'                str_Buff = Chr$(10)
'       Case TM_T88
'                 str_Buff = Chr$(10)
'       Case Else
'End Select
'Select Case rs!Port
'           Case "Rs232c"
'                        If (main.MSComm(TD_PORT).PortOpen = True) Then
'                            For i = 1 To Feed_Line
'                                main.MSComm(TD_PORT).Output = str_Buff
'                            Next i
'                        End If
'           Case "LPT1"
'                        For i = 1 To Feed_Line
'                            Rtn = WriteFileNOString(F_Handle, str_Buff, Len(str_Buff), Len(str_Buff), 0)
'                        Next i
'           Case "LPT2"
'                        For i = 1 To Feed_Line
'                            Rtn = WriteFileNOString(F_Handle, str_Buff, Len(str_Buff), Len(str_Buff), 0)
'                        Next i
'End Select
End Sub

Public Sub TDPrint_String(PrintString As String, mode As Byte, F_Handle As Long)
'Dim Rtn As Long
'Dim str_Buff As String
'Dim tmp As Byte
'Dim rs As Recordset
'
'Set rs = ParkDb.OpenRecordset("SELECT * FROM 영수증프린터", dbOpenSnapshot)
'If rs.EOF Then
'   Set rs = Nothing
'   Exit Sub
'End If
'
'With main
'tmp = Print_Type
'Select Case tmp
'       Case CBM_720
'                    If (Mode = 0) Then
'                        str_Buff = Chr$(15)
'                    Else
'                        str_Buff = Chr$(14)
'                    End If
'                    Select Case rs!Port
'                               Case "Rs232c"
'                                            If (.MSComm(TD_PORT).PortOpen = True) Then
'                                                .MSComm(TD_PORT).Output = str_Buff
'                                                .MSComm(TD_PORT).Output = PrintString & Chr$(10)
'                                            End If
'                               Case "LPT1"
'                                            Rtn = WriteFileNOString(F_Handle, str_Buff, Len(str_Buff), Len(str_Buff), 0)
'                                            Rtn = WriteFileNOString(F_Handle, PrintString & Chr$(10), LenH(PrintString & Chr$(10)), LenH(PrintString & Chr$(10)), 0)
'                               Case "LPT2"
'                                            Rtn = WriteFileNOString(F_Handle, str_Buff, Len(str_Buff), Len(str_Buff), 0)
'                                            Rtn = WriteFileNOString(F_Handle, PrintString & Chr$(10), LenH(PrintString & Chr$(10)), LenH(PrintString & Chr$(10)), 0)
'                               Case "FILE"
'                                            Call Print_doc(PrintString)
'                    End Select
'       Case CP_300
'                    If (Mode = 0) Then
'                        str_Buff = Chr$(20)
'                    Else
'                        str_Buff = Chr$(14)
'                    End If
'                    Select Case rs!Port
'                               Case "Rs232c"
'                                            If (.MSComm(TD_PORT).PortOpen = True) Then
'                                                .MSComm(TD_PORT).Output = str_Buff
'                                                .MSComm(TD_PORT).Output = PrintString & Chr$(10)
'                                            End If
'                               Case "LPT1"
'                                            Rtn = WriteFileNOString(F_Handle, str_Buff, Len(str_Buff), Len(str_Buff), 0)
'                                            Rtn = WriteFileNOString(F_Handle, PrintString & Chr$(10), LenH(PrintString & Chr$(10)), LenH(PrintString & Chr$(10)), 0)
'                               Case "LPT2"
'                                            Rtn = WriteFileNOString(F_Handle, str_Buff, Len(str_Buff), Len(str_Buff), 0)
'                                            Rtn = WriteFileNOString(F_Handle, PrintString & Chr$(10), LenH(PrintString & Chr$(10)), LenH(PrintString & Chr$(10)), 0)
'                               Case "FILE"
'                                            Call Print_doc(PrintString)
'                    End Select
'       Case TM_300
'                    If (Mode = 0) Then
'                        str_Buff = Chr$(&H1B) & Chr$(&H21) & Chr$(&H0)
'                    Else
'                        str_Buff = Chr$(&H1B) & Chr$(&H21) & Chr$(&H20)
'                    End If
'                    Select Case rs!Port
'                               Case "Rs232c"
'                                            If (.MSComm(TD_PORT).PortOpen = True) Then
'                                                .MSComm(TD_PORT).Output = str_Buff
'                                                .MSComm(TD_PORT).Output = PrintString & Chr$(10)
'                                            End If
'                               Case "LPT1"
'                                            Rtn = WriteFileNOString(F_Handle, str_Buff, Len(str_Buff), Len(str_Buff), 0)
'                                            Rtn = WriteFileNOString(F_Handle, PrintString & Chr$(10), LenH(PrintString & Chr$(10)), LenH(PrintString & Chr$(10)), 0)
'                               Case "LPT2"
'                                            Rtn = WriteFileNOString(F_Handle, str_Buff, Len(str_Buff), Len(str_Buff), 0)
'                                            Rtn = WriteFileNOString(F_Handle, PrintString & Chr$(10), LenH(PrintString & Chr$(10)), LenH(PrintString & Chr$(10)), 0)
'                               Case "FILE"
'                                            Call Print_doc(PrintString)
'                    End Select
'       Case STAR_300
'                    If (Mode = 0) Then
'                        str_Buff = Chr$(15)
'                    Else
'                        str_Buff = Chr$(14)
'                    End If
'                    Select Case rs!Port
'                               Case "Rs232c"
'                                            If (.MSComm(TD_PORT).PortOpen = True) Then
'                                                .MSComm(TD_PORT).Output = str_Buff
'                                                .MSComm(TD_PORT).Output = PrintString & Chr$(10)
'                                            End If
'                               Case "LPT1"
'                                            Rtn = WriteFileNOString(F_Handle, str_Buff, Len(str_Buff), Len(str_Buff), 0)
'                                            Rtn = WriteFileNOString(F_Handle, PrintString & Chr$(10), LenH(PrintString & Chr$(10)), LenH(PrintString & Chr$(10)), 0)
'                               Case "LPT2"
'                                            Rtn = WriteFileNOString(F_Handle, str_Buff, Len(str_Buff), Len(str_Buff), 0)
'                                            Rtn = WriteFileNOString(F_Handle, PrintString & Chr$(10), LenH(PrintString & Chr$(10)), LenH(PrintString & Chr$(10)), 0)
'                               Case "FILE"
'                                            Call Print_doc(PrintString)
'                    End Select
'       Case TM_T88
'                    If (Mode = 0) Then
'                        str_Buff = Chr$(&H1B) & Chr$(&H21) & Chr$(&H0)
'                    Else
'                        str_Buff = Chr$(&H1B) & Chr$(&H21) & Chr$(&H20)
'                    End If
'                    Select Case rs!Port
'                               Case "Rs232c"
'                                            If (.MSComm(TD_PORT).PortOpen = True) Then
'                                                .MSComm(TD_PORT).Output = str_Buff
'                                                .MSComm(TD_PORT).Output = PrintString & Chr$(10)
'                                            End If
'                               Case "LPT1"
'                                            Rtn = WriteFileNOString(F_Handle, str_Buff, Len(str_Buff), Len(str_Buff), 0)
'                                            Rtn = WriteFileNOString(F_Handle, PrintString & Chr$(10), LenH(PrintString & Chr$(10)), LenH(PrintString & Chr$(10)), 0)
'                               Case "LPT2"
'                                            Rtn = WriteFileNOString(F_Handle, str_Buff, Len(str_Buff), Len(str_Buff), 0)
'                                            Rtn = WriteFileNOString(F_Handle, PrintString & Chr$(10), LenH(PrintString & Chr$(10)), LenH(PrintString & Chr$(10)), 0)
'                               Case "FILE"
'                                            Call Print_doc(PrintString)
'                    End Select
'       Case Else
'End Select
'End With
End Sub



