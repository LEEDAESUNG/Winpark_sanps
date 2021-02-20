Attribute VB_Name = "modUnita"
Option Explicit


Public Sub Unita_Send(Index As Integer, cmd As String)

    Dim unitaIP As String
    Dim unitaPort As Long
    Dim unitaDeviceMode As String
    Dim Qry As String
    
On Error GoTo Err_P

    Select Case Index
        Case 0
            unitaIP = Trim(LANE1_UnitaIP)
            unitaPort = LANE1_UnitaPort
            unitaDeviceMode = LANE1_UnitaDeviceMode
        Case 1
            unitaIP = Trim(LANE2_UnitaIP)
            unitaPort = LANE2_UnitaPort
            unitaDeviceMode = LANE2_UnitaDeviceMode
        Case 2
            unitaIP = Trim(LANE3_UnitaIP)
            unitaPort = LANE3_UnitaPort
            unitaDeviceMode = LANE3_UnitaDeviceMode
        Case 3
            unitaIP = Trim(LANE4_UnitaIP)
            unitaPort = LANE4_UnitaPort
            unitaDeviceMode = LANE4_UnitaDeviceMode
        Case 4
            unitaIP = Trim(LANE5_UnitaIP)
            unitaPort = LANE5_UnitaPort
            unitaDeviceMode = LANE5_UnitaDeviceMode
        Case 5
            unitaIP = Trim(LANE6_UnitaIP)
            unitaPort = LANE6_UnitaPort
            unitaDeviceMode = LANE6_UnitaDeviceMode
    End Select
    
    
    
    Unita_Str(Index) = cmd

    Select Case unitaDeviceMode
        Case "0"    ' TCP
            'If (glo_check = False) Then
            
                If (FrmTcpServer.Unita1_sock(Index).State <> sckClosed) Then
                    FrmTcpServer.Unita1_sock(Index).Close
                End If
    
                FrmTcpServer.Unita1_sock(Index).Connect Trim(unitaIP), unitaPort
                Call None_Delay_Time(0.1)
                
                
                'glo_check = True
                
                
            'Else
            '    FrmTcpServer.Unita1_sock(Index).SendData Unita_Str(Index)
            'End If
            
            
            
            Qry = "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('', 'HOST','Unita" & Index + 1 & " TCP Send: " & Unita_Str(Index) & "',''," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
            adoConn.Execute Qry
            Call DataLogger("[Unita TCP Send]  " & cmd) '임시테스트
        Case "1"    ' UDP
            FrmTcpServer.Unita1_sock(Index).SendData Unita_Str(Index)
            
            Qry = "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('', 'HOST','Unita" & Index + 1 & " UDP Send: " & Unita_Str(Index) & "',''," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
            adoConn.Execute Qry
            Call DataLogger("[Unita UDP Send]  " & cmd) '임시테스트
    End Select
    
    Exit Sub
    
Err_P:
    If unitaDeviceMode = 0 Then
        Call DataLogger("Unita Send Err:" & Err.Description & "Index:" & Index & ", Protocol:TCP, Connect IP:" & unitaIP & "Port:" & CStr(unitaPort) & ", cmd:" & cmd)
    Else
        Call DataLogger("Unita Send Err:" & Err.Description & "Index:" & Index & ", Protocol:UDP, cmd:" & cmd)
    End If
End Sub



Public Sub Unita_Command_Send(Index As Integer, cmd As String)

    Dim unitaIP As String
    Dim unitaPort As Long
    Dim unitaDeviceMode As String
    Dim Qry As String
    
On Error GoTo Err_P

    Select Case Index
        Case 0
            unitaIP = Trim(LANE1_UnitaIP)
            unitaPort = LANE1_UnitaPort
            unitaDeviceMode = LANE1_UnitaDeviceMode
        Case 1
            unitaIP = Trim(LANE2_UnitaIP)
            unitaPort = LANE2_UnitaPort
            unitaDeviceMode = LANE2_UnitaDeviceMode
        Case 2
            unitaIP = Trim(LANE3_UnitaIP)
            unitaPort = LANE3_UnitaPort
            unitaDeviceMode = LANE3_UnitaDeviceMode
        Case 3
            unitaIP = Trim(LANE4_UnitaIP)
            unitaPort = LANE4_UnitaPort
            unitaDeviceMode = LANE4_UnitaDeviceMode
        Case 4
            unitaIP = Trim(LANE5_UnitaIP)
            unitaPort = LANE5_UnitaPort
            unitaDeviceMode = LANE5_UnitaDeviceMode
        Case 5
            unitaIP = Trim(LANE6_UnitaIP)
            unitaPort = LANE6_UnitaPort
            unitaDeviceMode = LANE6_UnitaDeviceMode
    End Select
    
    
    
    Unita_Cmd_Str(Index) = cmd

    Select Case unitaDeviceMode
        Case "0"    ' TCP
            If (FrmTcpServer.Unita1_cmd_sock(Index).State <> sckClosed) Then
                FrmTcpServer.Unita1_cmd_sock(Index).Close
            End If
            FrmTcpServer.Unita1_cmd_sock(Index).Connect ' Trim(unitaIP), unitaPort
            
            Qry = "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('', 'HOST','Unita" & Index + 1 & " TCP Send: " & Unita_Cmd_Str(Index) & "',''," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
            adoConn.Execute Qry
            Call DataLogger("[Unita TCP Send]  " & cmd) '임시테스트
        Case "1"    ' UDP
            FrmTcpServer.Unita1_cmd_sock(Index).SendData Unita_Cmd_Str(Index)
            
            Qry = "INSERT INTO tb_log(TICKET_CODE, PROC_CODE, PROC_INFO, ACCOUNT_NAME, ACCOUNT_MONEY, REG_DATE ) VALUES ('', 'HOST','Unita" & Index + 1 & " UDP Send: " & Unita_Cmd_Str(Index) & "',''," & 0 & ",'" & Format(Now, "yyyy-mm-dd hh:nn:ss") & "')"
            adoConn.Execute Qry
            Call DataLogger("[Unita UDP Send]  " & cmd) '임시테스트
    End Select
    
    Exit Sub
    
Err_P:
    If unitaDeviceMode = 0 Then
        Call DataLogger("Unita Cmd Send Err:" & Err.Description & "Index:" & Index & ", Protocol:TCP, Connect IP:" & unitaIP & "Port:" & CStr(unitaPort) & ", cmd:" & cmd)
    Else
        Call DataLogger("Unita Cmd Send Err:" & Err.Description & "Index:" & Index & ", Protocol:UDP, cmd:" & cmd)
    End If
End Sub


