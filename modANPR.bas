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
    
    Result = Rec_EngineResultPB(False, fileName, ByteResult(0), Rec, CRec(0), Conf(0), 0, 0, 0, 0)          'VB6.0���� ��� : SRec���� ���������� ������.
    'Result = Rec_EngineResultEx(False, fileName, ByteResult(0), Rec, CRec(0), Conf(0), SRec, 0, 0, 0, 0)   'VB6.0�̻󿡼� ���
        
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
    
    SRec.Left = 0           ' �ν��ϰ��� �ϴ� �̹����� ���� ��� 0�϶� ��üȭ������ ������
    SRec.Top = 0
    SRec.Right = 0
    SRec.Bottom = 0
    
    'img(0) : ���� �����̵���Ÿ (�����̹����� 8bit �׷��� �̹����� ������)
    'width : BitPerLine (���λ�����, 4�� �����)
    'bpp : BitsPerPixel (8bit)
    'width : ����ũ��
    'height : ����ũ��
    'ByteResult(0) : �νİ��
    'Rec : �ν��� ��ȣ�� ��ġ
    'CRec(0) : ��ȣ�� ���� ���� ��ġ. �ּ� 16���� �Ҵ��Ͻÿ�.
    'Conf(0) : ��ȣ�� ���� �ŷڵ�. �ּ� 16���� �Ҵ��Ͻÿ�. ����� �ŷڵ��� �� �����ʾƼ� ���� �Ѿ���� ����.
    'SRec : ��ȣ���� �ִ� �������� ��ġ. ��ȣ���� �� �簢�� �ȿ� �־�� ��. (0 �Ͻ� ��ü)
    '0, 0 : ����, ���� �̹��� ���� ���� ( �̹����� Ȯ��, ����ؼ� �ν���, �̶� ��ȣ�� ��ġ���� Ȯ��,��ҵ� ������ ����
    '0, 0 : ����, ���� ������ ���� ���� * �ڼ��� ������ ÷�ε� ���� ����..
    
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


