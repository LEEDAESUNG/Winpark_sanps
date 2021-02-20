Attribute VB_Name = "modANPR"
Option Explicit


Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long



Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type



Public Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * 100 'MAX_PATH
        cAlternate As String * 14
End Type



Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long



















Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Declare Function Rec_EngineClose Lib "Rec_Module.dll" () As Long
Public Declare Function Rec_EngineOpen Lib "Rec_Module.dll" () As Long
Public Declare Function Rec_EngineResultEx Lib "Rec_Module.dll" (ByVal bBuffer As Boolean, ByVal ImageFileName As String, _
                                                                        ByRef Recog_Str As Byte, ByRef Rec As RECT, ByRef CRec As RECT, _
                                                                        ByRef Conf As Integer, ByVal SRec As RECT, ByVal rat_x As Integer, _
                                                                        ByVal rat_y As Integer, ByVal deg_x As Integer, _
                                                                        ByVal deg_y As Integer) As Long
                                                                        
Public Declare Function Rec_EngineResultPB Lib "Rec_Module.dll" (ByVal bBuffer As Boolean, ByVal ImageFileName As String, _
                                                                        ByRef Recog_Str As Byte, ByRef Rec As RECT, ByRef CRec As RECT, _
                                                                        ByRef Conf As Integer, ByVal rat_x As Integer, _
                                                                        ByVal rat_y As Integer, ByVal deg_x As Integer, _
                                                                        ByVal deg_y As Integer) As Long
                                                                        
Public Declare Function Rec_EngineResultPB_Buf Lib "Rec_Module.dll" (Imagebuf As Any, ByVal bpl As Integer, ByVal bpp As Integer, _
                                                                           ByVal size_w As Integer, ByVal size_h As Integer, _
                                                                           ByRef Recog_Str As Byte, ByRef Rec As RECT, ByRef CRec As RECT, _
                                                                           ByRef Conf As Integer, ByRef SRec As RECT, ByVal rat_x As Integer, _
                                                                           ByVal rat_y As Integer, ByVal deg_x As Integer, _
                                                                           ByVal deg_y As Integer) As Long

Public Function GetPlateNumber(fileName As String) As String
    Dim Result As Long
    Dim ByteResult(50) As Byte
    Dim Str_Result As String
    
    Dim Rec As RECT
    Dim CRec(16) As RECT
    Dim SRec As RECT
    
    Dim Conf(16) As Integer
    
    ChDir App.Path
    
    SRec.Top = 0
    SRec.Left = 0
    SRec.Right = 0
    SRec.Bottom = 0
    
    Result = Rec_EngineResultPB(False, fileName, ByteResult(0), Rec, CRec(0), Conf(0), 0, 0, 0, 0)          'VB6.0에서 사용 : SRec변수 에러로인해 사용안함.
    'Result = Rec_EngineResultEx(False, fileName, ByteResult(0), Rec, CRec(0), Conf(0), SRec, 0, 0, 0, 0)   'VB6.0이상에서 사용
        
    Str_Result = StrConv(ByteResult, vbUnicode)
    
    If Len(Str_Result) <= 4 Then
        Result = -1
    End If
    
       
    If Result = -1 Then
        GetPlateNumber = "XXXXXXX"
    Else
        GetPlateNumber = Str_Result
    End If
End Function

Public Function GetPlateNumber_buf(img() As Byte, width As Long, height As Long, bpp As Long) As String
    Dim Result As Long
    Dim ByteResult(50) As Byte
    Dim Str_Result As String
    
    Dim Rec As RECT
    Dim CRec(16) As RECT
    Dim SRec As RECT
    
    Dim Conf(16) As Integer
    
    ChDir App.Path
    
    SRec.Left = 0           ' 인식하고자 하는 이미지의 영역 모두 0일때 전체화면으로 설정됨
    SRec.Top = 0
    SRec.Right = 0
    SRec.Bottom = 0
    
    'img(0) : 버퍼 이지미데이타 (버퍼이미지는 8bit 그레이 이미지만 지원함)
    'width : BitPerLine (가로사이즈, 4의 배수인)
    'bpp : BitsPerPixel (8bit)
    'width : 가로크기
    'height : 세로크기
    'ByteResult(0) : 인식결과
    'Rec : 인식한 번호판 위치
    'CRec(0) : 번호판 내의 숫자 위치. 최소 16개를 할당하시오.
    'Conf(0) : 번호판 숫의 신뢰도. 최소 16개를 할당하시오. 현재는 신뢰도가 잘 맞지않아서 값이 넘어오지 않음.
    'SRec : 번호판이 있는 개략적인 위치. 번호판은 이 사각형 안에 있어야 함. (0 일시 전체)
    '0, 0 : 가로, 세로 이미지 비율 설정 ( 이미지를 확대, 축소해서 인식함, 이때 번호판 위치값이 확대,축소된 값으로 나옴
    '0, 0 : 가로, 세로 기울어진 각도 설정 * 자세한 내용은 첨부된 설명서 참조..
    
    Result = Rec_EngineResultPB_Buf(img(0), width, bpp, width, height, ByteResult(0), Rec, CRec(0), Conf(0), SRec, 0, 0, 0, 0)
        
    Str_Result = StrConv(ByteResult, vbUnicode)
    
    Dim L, T, R, B As String
    
    L = Rec.Left
    T = Rec.Top
    R = Rec.Right
    B = Rec.Bottom
    
    If Len(Str_Result) <= 4 Then
        Result = -1
    End If
       
    If Result = -1 Then
        GetPlateNumber_buf = "XXXXXXX"
    Else
        GetPlateNumber_buf = Str_Result
    End If
    
End Function


