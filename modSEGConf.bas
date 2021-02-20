Attribute VB_Name = "modSEGConf"
Option Explicit

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' WIZ120SR Board configuration data type
'--------------------------------------------------
Private Type typeBoardInfo
'--------------------------------------------------
' Common data
'--------------------------------------------------
    mac(0 To 5) As Byte
    ip(0 To 3) As Byte
    subnet(0 To 3) As Byte
    gw(0 To 3) As Byte
    DHCP As Byte
    AppVer(0 To 1) As Byte
    debugoff As Byte
'--------------------------------------------------
' each channel data
'--------------------------------------------------
' UART 0
    bserver0 As Byte
    myport0(0 To 1) As Byte
    peerip0(0 To 3) As Byte
    peerport0(0 To 1) As Byte
    speed0 As Byte
    databit0 As Byte
    parity0 As Byte
    stopbit0 As Byte
    flow0 As Byte
    D_ch0 As Byte
    D_size0(0 To 1) As Byte
    D_time0(0 To 1) As Byte
    I_time0(0 To 1) As Byte
    UDP0 As Byte
    Connect0 As Byte
    DNS_Flag0 As Byte
    DNS_IP0(0 To 3) As Byte
    D_SIP0(0 To 31) As Byte
    EnTCPPass0 As Byte
    TCPPass0(0 To 7) As Byte
' UART 1
    bserver1 As Byte
    myport1(0 To 1) As Byte
    peerip1(0 To 3) As Byte
    peerport1(0 To 1) As Byte
    speed1 As Byte
    databit1 As Byte
    parity1 As Byte
    stopbit1 As Byte
    flow1 As Byte
    D_ch1 As Byte
    D_size1(0 To 1) As Byte
    D_time1(0 To 1) As Byte
    I_time1(0 To 1) As Byte
    UDP1 As Byte
    Connect1 As Byte
    DNS_Flag1 As Byte
    DNS_IP1(0 To 3) As Byte
    D_SIP1(0 To 31) As Byte
    EnTCPPass1 As Byte
    TCPPass1(0 To 7) As Byte
'--------------------------------------------------
' Common data
'--------------------------------------------------
    SCfg0 As Byte
    SCfgStr0(0 To 2) As Byte
    PPPoE_ID(0 To 31) As Byte
    PPPoE_Pass(0 To 31) As Byte
'--------------------------------------------------
End Type

Public Const BoardInfoSize_3_4 As Integer = 232


' This tool's mode
'--------------------------------------------------
Public Enum typeToolMode
    modeNone = 0
    modeSearching = 1
    modeSetting = 2
    modeSettingComplete = 3
    modeUploading = 4
    modeUploadingComplete = 5
End Enum
Public ToolMode As typeToolMode

' Total count of Boards
'--------------------------------------------------
Public intBoardNum As Integer
' Collection of Board's configuration data.
'--------------------------------------------------
Public colBoards As New Collection

' Selected Board's infomation
'--------------------------------------------------
Public BoardKey As String
Public BoardInfo As typeBoardInfo
Public bSelect As Boolean
Public bDirectUpload As Boolean

' Selected Firmware file and DestIP address for uploading
'--------------------------------------------------
Public strUploadFile As String
Public destIP As String

Sub MessageBox(msg As String)
       
    Call MsgBox(msg, vbInformation Or vbMsgBoxSetForeground, "WIZnet SEG Information")
    
End Sub
