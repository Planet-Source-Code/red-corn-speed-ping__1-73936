Attribute VB_Name = "modCopyMemory"
Option Explicit
Public Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long
End Type

Public Const WM_COPYDATA = &H4A

Public Const MSG_MYPING_BASE As Long = 814692
Public Const MSG_AGENT_LOADPINGLIST As Long = MSG_MYPING_BASE + 1
Public Const MSG_AGENT_SAYHELLO As Long = MSG_MYPING_BASE + 2
Public Const MSG_AGENT_STARTPING As Long = MSG_MYPING_BASE + 3
Public Const MSG_AGENT_STOPPING As Long = MSG_MYPING_BASE + 4
Public Const MSG_REPORT_PINGSTATUS_RTT As Long = MSG_MYPING_BASE + 5
Public Const MSG_REPORT_PINGSTATUS_RECV As Long = MSG_MYPING_BASE + 6
Public Const MSG_REPORT_PINGSTATUS_LOST As Long = MSG_MYPING_BASE + 7
Public Const MSG_REPORT_PINGSTATUS_OK As Long = MSG_MYPING_BASE + 8
Public Const MSG_REPORT_PINGSTATUS_STATUS As Long = MSG_MYPING_BASE + 9
Public Const MSG_AGENT_READY As Long = MSG_MYPING_BASE + 10
Public Const MSG_AGENT_CLOSE As Long = MSG_MYPING_BASE + 11
Public Const MSG_REPORT_PINGSTATUS_INFO As Long = MSG_MYPING_BASE + 12
Public Const MSG_AGENT_BLINK As Long = MSG_MYPING_BASE + 13
Public Const MSG_AGENT_SAYGOODBYE As Long = MSG_MYPING_BASE + 14
Public Const MSG_AGENT_SET_CYCLEINTERVAL As Long = MSG_MYPING_BASE + 15
Public Const MSG_AGENT_LOAD_INI As Long = MSG_MYPING_BASE + 16
Public Const MSG_AGENT_REPORT_LOAD_INI_OK As Long = MSG_MYPING_BASE + 17
Public Const MSG_PLS_SAYHELLO As Long = MSG_MYPING_BASE + 18
Public Const MSG_REPORT_PINGSTAT As Long = MSG_MYPING_BASE + 19
Public Const MSG_AGENT_KEEPALIVE As Long = MSG_MYPING_BASE + 20
Public Const MSG_AGENT_BLINK_A As Long = MSG_MYPING_BASE + 21
Public Const MSG_AGENT_BLINK_B As Long = MSG_MYPING_BASE + 22
Public Const MSG_AGENT_BLINK_S As Long = MSG_MYPING_BASE + 23

Public Const MSG_LOGAGENT_SAYHELLO As Long = MSG_MYPING_BASE + 101
Public Const MSG_LOGAGENT_CLOSE As Long = MSG_MYPING_BASE + 102
Public Const MSG_LOGAGENT_SAYGOODBYE As Long = MSG_MYPING_BASE + 103

Public Const MSG_LOGAGENT_LOADPINGLIST As Long = MSG_MYPING_BASE + 105
Public Const MSG_LOGAGENT_REPORT_READY As Long = MSG_MYPING_BASE + 106
Public Const MSG_LOGAGENT_REPORT_LOAD_INI_OK As Long = MSG_MYPING_BASE + 107

Public Const MSG_LOGAGENT_LOGSTAT_BEGIN As Long = MSG_MYPING_BASE + 109
Public Const MSG_LOGAGENT_LOGSTAT_DATA As Long = MSG_MYPING_BASE + 110
Public Const MSG_LOGAGENT_LOGSTAT_END As Long = MSG_MYPING_BASE + 111

Public Const MSG_LOGAGENT_BLINK As Long = MSG_MYPING_BASE + 112

Public Const MSG_LOGAGENT_SAVELOG_START As Long = MSG_MYPING_BASE + 113
Public Const MSG_LOGAGENT_SAVELOG_BUF1 As Long = MSG_MYPING_BASE + 114
Public Const MSG_LOGAGENT_SAVELOG_BUF2 As Long = MSG_MYPING_BASE + 115
Public Const MSG_LOGAGENT_SAVELOG_END As Long = MSG_MYPING_BASE + 116

Public Const MSG_PINGCOLLECTOR_SAYHELLO As Long = MSG_MYPING_BASE + 201
Public Const MSG_PLS_PINGCOLLECTOR_MOVE As Long = MSG_MYPING_BASE + 202
Public Const MSG_PLS_PINGCOLLECTOR_CLOSE As Long = MSG_MYPING_BASE + 203
Public Const MSG_PLS_PINGCOLLECTOR_LOADPINGLIST As Long = MSG_MYPING_BASE + 204
Public Const MSG_PLS_PINGCOLLECTOR_CREATEAGENT As Long = MSG_MYPING_BASE + 205
Public Const MSG_PLS_PINGCOLLECTOR_LOADINI As Long = MSG_MYPING_BASE + 206
Public Const MSG_PINGCOLLECTOR_REPORT_LOADINI_OK As Long = MSG_MYPING_BASE + 207
Public Const MSG_PINGCOLLECTOR_READY As Long = MSG_MYPING_BASE + 208
Public Const MSG_PINGCOLLECTOR_SAYGOODBYE As Long = MSG_MYPING_BASE + 209
Public Const MSG_PINGCOLLECTOR_STARTPING As Long = MSG_MYPING_BASE + 210
Public Const MSG_PINGCOLLECTOR_STOPPING As Long = MSG_MYPING_BASE + 211
Public Const MSG_PINGCOLLECTOR_LOADPINGLIST_ERR As Long = MSG_MYPING_BASE + 212
Public Const MSG_PINGCOLLECTOR_LOADPINGLIST_OK As Long = MSG_MYPING_BASE + 213
Public Const MSG_PLS_PINGCOLLECTOR_REPORT_PING_RESULT As Long = MSG_MYPING_BASE + 214
Public Const MSG_PINGCOLLECTOR_REPORT_PING_RESULT As Long = MSG_MYPING_BASE + 215
Public Const MSG_PINGCOLLECTOR_BLINK_A As Long = MSG_MYPING_BASE + 216
Public Const MSG_PINGCOLLECTOR_BLINK_B As Long = MSG_MYPING_BASE + 217
Public Const MSG_PINGCOLLECTOR_BLINK_S As Long = MSG_MYPING_BASE + 218
Public Const MSG_PLS_PINGCOLLECTOR_SEND_EVENT As Long = MSG_MYPING_BASE + 219

Public Const MSG_PINGCOLLECTOR_SENDEVENT_BUF1 As Long = MSG_MYPING_BASE + 221
Public Const MSG_PINGCOLLECTOR_SENDEVENT_BUF2 As Long = MSG_MYPING_BASE + 222
Public Const MSG_PINGCOLLECTOR_SENDEVENT_NOEVENT As Long = MSG_MYPING_BASE + 223

Public Const MSG_MYPING_POSTMSG As String = "MsgMyPing1.0"

Public Const MSG_KEEPALIVE_REQUEST As Long = MSG_MYPING_BASE + 400
Public Const MSG_PINGAGENT_KEEPALIVE_REPLY As Long = MSG_MYPING_BASE + 401
Public Const MSG_PINGCOLLECTOR_KEEPALIVE_REPLY As Long = MSG_MYPING_BASE + 402
Public Const MSG_STOP_PING As Long = MSG_MYPING_BASE + 403





'Copies a block of memory from one location to another.
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReplyMessage Lib "user32" (ByVal lReply As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


'//Ascii to Unicode

'The WideCharToMultiByte function maps a wide-character string to a new character string.
'The function is faster when both lpDefaultChar and lpUsedDefaultChar are NULL.

'CodePage
Private Const CP_ACP = 0 'ANSI
Private Const CP_MACCP = 2 'Mac
Private Const CP_OEMCP = 1 'OEM
Private Const CP_UTF7 = 65000
Private Const CP_UTF8 = 65001

'dwFlags
Private Const WC_NO_BEST_FIT_CHARS = &H400
Private Const WC_COMPOSITECHECK = &H200
Private Const WC_DISCARDNS = &H10
Private Const WC_SEPCHARS = &H20 'Default
Private Const WC_DEFAULTCHAR = &H40

Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, _
                                                    ByVal dwFlags As Long, _
                                                    ByVal lpWideCharStr As Long, _
                                                    ByVal cchWideChar As Long, _
                                                    ByVal lpMultiByteStr As Long, _
                                                    ByVal cbMultiByte As Long, _
                                                    ByVal lpDefaultChar As Long, _
                                                    ByVal lpUsedDefaultChar As Long) As Long
                                                    
'//subclassing
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const WM_CLOSE As Long = &H10
Private Const GWL_WNDPROC = (-4)
Dim LocalPrevWndProc As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long
    
Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Sub Hook()
  On Error Resume Next
  LocalPrevWndProc = SetWindowLong(glMyHwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub
Public Sub UnHook()
  Dim WorkFlag As Long

  On Error Resume Next
  WorkFlag = SetWindowLong(glMyHwnd, GWL_WNDPROC, LocalPrevWndProc)
End Sub

Private Function WindowProc(ByVal Lwnd As Long, ByVal Lmsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'        lParam: pointer to structure with data
'        wParam: handle of sending window
    Dim cdCopyData As COPYDATASTRUCT
    Dim LenOfCopyData As Long
    Dim SN As Long
    Dim bufAgentSayHello(1 To 2) As Long
    Dim AgentID As Long
    Dim AgentHwnd As Long
    Dim i As Integer
    Dim bufBlinkInfo(1 To 2) As Long
    'Dim BlinkSwitch As Integer
    Dim LogBlinkSwitch As Integer
    Dim rStart As Long, rEnd As Long
    Dim NumOfLogList As Long
    Dim num As Long
    'Dim tmpAryLost() As Integer
    
    Select Case Lmsg
    Case glMyPingMsg
        Select Case wParam
        Case MSG_AGENT_BLINK_A
            frmMain.AgentLedBlink lParam, MSG_AGENT_BLINK_A
        Case MSG_AGENT_BLINK_B
            frmMain.AgentLedBlink lParam, MSG_AGENT_BLINK_B
        Case MSG_AGENT_BLINK_S
            frmMain.AgentLedBlink lParam, MSG_AGENT_BLINK_S
        Case MSG_AGENT_SAYGOODBYE 'Say Goodbye®É¥Îpostmessage,Â÷¶}ªº°®¯Ü¤@ÂI
            AgentID = lParam
            frmMain.AgentSayGoodbye AgentID
        Case MSG_PLS_PINGCOLLECTOR_REPORT_PING_RESULT
            ReportPingResult
        Case MSG_PLS_PINGCOLLECTOR_SEND_EVENT
            SendEventToManager
        Case MSG_PLS_PINGCOLLECTOR_CLOSE
            frmMain.CloseMe
        End Select
            
    Case WM_COPYDATA
        
        Call CopyMemory(cdCopyData, ByVal lParam, Len(cdCopyData))
        'MsgBox cdCopyData.dwData
        
        LenOfCopyData = cdCopyData.cbData
        Select Case cdCopyData.dwData
            Case MSG_PINGCOLLECTOR_STARTPING
                ReplyMessage 1
                frmMain.StartPing
            Case MSG_STOP_PING
                ReplyMessage 1
                frmMain.StopPing
            Case MSG_PLS_PINGCOLLECTOR_LOADINI
                LenOfCopyData = cdCopyData.cbData
                Call CopyMemory(aryINI(1), ByVal cdCopyData.lpData, LenOfCopyData)
                ReplyMessage 1

                glPingCount = aryINI(1)
                glPingTimeOutHost = aryINI(2)
                glPingMaxBurst = aryINI(3)
                glPingTimeOutBatch = aryINI(4)
                glWaitForSingleObject = aryINI(5)
                glCycleInterval = aryINI(6)
                
                glAgentCount = aryINI(7)
                glThreshold = aryINI(8)
                glContinuedFailAsDown = aryINI(9)
                glStatisticsCycle = aryINI(10)
                glRefreshCycle = aryINI(11)
                glDelayStart = aryINI(12)
                
                If glPingCount > 0 And _
                    glPingTimeOutHost > 0 And _
                    glPingMaxBurst > 0 And _
                    glPingTimeOutBatch > 0 And _
                    glWaitForSingleObject > 0 And _
                    glCycleInterval > 0 And _
                    glAgentCount > 0 And _
                    glThreshold > 0 And _
                    glContinuedFailAsDown > 0 And _
                    glStatisticsCycle > 0 And _
                    glRefreshCycle > 0 And _
                    glDelayStart > 0 Then
                    frmMain.InitAgentStuff '¥ýªì©l¤ÆAgent Array, pingnode ub & lb...
                    'ReportLoadIniOK
                    Call ReportReady '§ï¦¨¦^³ø¦Û¤vready§Y¥i,agent¬O§_ready,¥Ñmanager¦Û¤v³B²z
                Else
                    MsgBox "¨t²Î°Ñ¼Æ¿ù»~!", vbExclamation, MsgTitle
                   
                End If
            
            Case MSG_PLS_PINGCOLLECTOR_LOADPINGLIST
                Call CopyMemory(num, ByVal cdCopyData.lpData, LenOfCopyData)
                ReplyMessage 1
                frmMain.DoLoadPingList num '--> ±N¦^³øLoadList OK / ERR
            
            Case MSG_PLS_PINGCOLLECTOR_MOVE
                ReplyMessage 1
                frmMain.MoveMyPos
'            Case MSG_AGENT_BLINK
'                Call CopyMemory(bufBlinkInfo(1), ByVal cdCopyData.lpData, LenOfCopyData)
'                ReplyMessage 1
'                'If AgentID > 0 And AgentID < 11 Then '¨¾¤î¿ù»~
'                    frmMain.AgentLedBlink bufBlinkInfo(1), bufBlinkInfo(2)
'                'End If
            
            Case MSG_REPORT_PINGSTATUS_INFO
                AgentID = wParam
                Call CopyMemory(aryAgentPingInfo(1, AgentID), ByVal cdCopyData.lpData, LenOfCopyData)
                aryAgentPingInfo(6, AgentID) = 1 'counter = 0
                aryAgentPingInfo(2, AgentID) = aryAgentLB(AgentID) + aryAgentPingInfo(2, AgentID)
                aryAgentPingInfo(3, AgentID) = aryAgentLB(AgentID) + aryAgentPingInfo(3, AgentID)
                ReplyMessage 1
            Case MSG_REPORT_PINGSTAT
                AgentID = wParam
                If aryAgentPingInfo(6, AgentID) <> 1 Then
                    ReplyMessage 1
                    frmMain.StopPing
                    MsgBox "MSG_REPORT_PINGSTATUS_RTT counter¿ù»~! ping°±¤î!"
                    Exit Function
                End If
                Call CopyMemory(aryAgentPingResultData(1, aryAgentPingInfo(2, AgentID)), ByVal cdCopyData.lpData, LenOfCopyData)
                aryAgentPingInfo(6, AgentID) = 2
                ReplyMessage 1
            Case MSG_REPORT_PINGSTATUS_OK
                '¦¬¨ì³Ì«áªº¦^³ø,­n¶}©l­pºâstatus
                AgentID = wParam
                If aryAgentPingInfo(6, AgentID) <> 2 Then
                    ReplyMessage 1
                    frmMain.StopPing
                    MsgBox "MSG_REPORT_PINGSTATUS_OK counter¿ù»~! ping°±¤î!"
                    Exit Function
                End If
                Call CopyMemory(SN, ByVal cdCopyData.lpData, LenOfCopyData)
                aryAgentPingInfo(6, AgentID) = 0
                ReplyMessage 1
                If aryAgentPingInfo(4, AgentID) <> SN Then
                    ReplyMessage 1
                    frmMain.StopPing
                    MsgBox "SN¿ù»~! ping°±¤î!"
                    Exit Function
                End If
                '·sª©ªº¤w¤£¦A§PÂ_pingcount¤F,¦]¬°pingcount = glpingcount®É¤~·|¦¬¨ì°T®§
                'MsgBox "agentid=" & AgentId & ", aryAgentPingInfo(3, AgentId)=" & aryAgentPingInfo(3, AgentId)
                frmMain.UpdatePingStatus AgentID, aryAgentPingInfo(2, AgentID), aryAgentPingInfo(3, AgentID), aryAgentPingInfo(5, AgentID)
        
            Case MSG_AGENT_SAYHELLO '·|±o¨ìagentid & agenthwnd
                Call CopyMemory(AgentHwnd, ByVal cdCopyData.lpData, LenOfCopyData)
                ReplyMessage 1
                AgentID = wParam
                aryAgentHwnd(AgentID) = AgentHwnd
                frmMain.AgentSayHello AgentID
            
            Case MSG_AGENT_REPORT_LOAD_INI_OK
                ReplyMessage 1
                AgentID = wParam
                'Agent Load List§¹¤§«á·|¦^³øReady
'                MsgBox "agentid=" & AgentID
'                MsgBox "aryAgentLB(AgentID)=" & aryAgentLB(AgentID)
'                MsgBox "aryAgentUB(AgentID)=" & aryAgentUB(AgentID)
'                MsgBox "aryAgentHwnd(AgentID)=" & aryAgentHwnd(AgentID)
                TellAgentToLoadPingList aryAgentLB(AgentID), aryAgentUB(AgentID), aryAgentHwnd(AgentID)
                
            Case MSG_AGENT_READY 'Agent Load List§¹¤§«á·|¦^³øReady
                ReplyMessage 1
                AgentID = wParam
                frmMain.AgentReady AgentID '-->§i¶DManager,§Ú¤]Ready¤F
        End Select
'        'Call InterProcessComms(lParam)
    End Select
    WindowProc = CallWindowProc(LocalPrevWndProc, Lwnd, Lmsg, wParam, lParam)
End Function

Public Sub SayGoodbye()
    If glPingManagerHwnd <> 0 Then
        PostMessage glPingManagerHwnd, glMyPingMsg, MSG_PINGCOLLECTOR_SAYGOODBYE, 0
    End If
End Sub
Public Sub ReportLoadListOK()
    Dim cdCopyData As COPYDATASTRUCT
    Dim DoCmd As Long
    If glPingManagerHwnd = 0 Then Exit Sub
    DoCmd = MSG_PINGCOLLECTOR_LOADPINGLIST_OK
    cdCopyData.dwData = DoCmd '¦Û©w¸q¼Æ¾Ú
    cdCopyData.cbData = Len(DoCmd)
    cdCopyData.lpData = VarPtr(DoCmd)
    SendMessage glPingManagerHwnd, WM_COPYDATA, glMyHwnd, cdCopyData
End Sub
Public Sub ReportLoadListERR()
    Dim cdCopyData As COPYDATASTRUCT
    Dim DoCmd As Long
    If glPingManagerHwnd = 0 Then Exit Sub
    DoCmd = MSG_PINGCOLLECTOR_LOADPINGLIST_ERR
    cdCopyData.dwData = DoCmd '¦Û©w¸q¼Æ¾Ú
    cdCopyData.cbData = Len(DoCmd)
    cdCopyData.lpData = VarPtr(DoCmd)
    SendMessage glPingManagerHwnd, WM_COPYDATA, glMyHwnd, cdCopyData
End Sub

Public Sub ReportReady()
    Dim cdCopyData As COPYDATASTRUCT
    Dim DoCmd As Long
    DoCmd = MSG_PINGCOLLECTOR_READY
    cdCopyData.dwData = DoCmd '¦Û©w¸q¼Æ¾Ú
    cdCopyData.cbData = Len(DoCmd)
    cdCopyData.lpData = VarPtr(DoCmd)
    SendMessage glPingManagerHwnd, WM_COPYDATA, glMyHwnd, cdCopyData
    '¦breadyªº®É­Ômanager¶}©lÀË¬dalive,¬G­n±Ò°Êblink
    frmMain.StartBlink
End Sub

Public Sub TellAgentToLoadPingList(lb As Long, ub As Long, AgentHwnd As Long)
    Dim cdCopyData As COPYDATASTRUCT
    Dim num As Long
    num = ub - lb + 1
    cdCopyData.dwData = MSG_AGENT_LOADPINGLIST '¦Û©w¸q¼Æ¾Ú
    cdCopyData.cbData = Len(aryInetAddr(0)) * num
    cdCopyData.lpData = VarPtr(aryInetAddr(lb))
    SendMessage AgentHwnd, WM_COPYDATA, glMyHwnd, cdCopyData
   
End Sub
Public Sub TellAgentToDoSomething(cmd As Long, AgentHwnd As Long)
    Dim cdCopyData As COPYDATASTRUCT
    Dim DoCmd As Long
    DoCmd = cmd
    cdCopyData.dwData = DoCmd '¦Û©w¸q¼Æ¾Ú
    
    cdCopyData.cbData = Len(DoCmd)
    cdCopyData.lpData = VarPtr(DoCmd)
    SendMessage AgentHwnd, WM_COPYDATA, glMyHwnd, cdCopyData
End Sub

Public Sub AgentSetCycleInterval(AgentHwnd As Long, CycleInterval As Long)
    Dim cdCopyData As COPYDATASTRUCT
    
    cdCopyData.dwData = MSG_AGENT_SET_CYCLEINTERVAL '¦Û©w¸q¼Æ¾Ú
    cdCopyData.cbData = Len(CycleInterval)
    cdCopyData.lpData = VarPtr(CycleInterval)
    SendMessage AgentHwnd, WM_COPYDATA, glMyHwnd, cdCopyData
End Sub
Public Sub CopyMemoryTest(cmd As Long, AgentHwnd As Long)
    Dim cdCopyData As COPYDATASTRUCT
    Dim DoCmd As Long
    DoCmd = cmd
    cdCopyData.dwData = DoCmd '¦Û©w¸q¼Æ¾Ú
    cdCopyData.cbData = Len(DoCmd)
    cdCopyData.lpData = VarPtr(DoCmd)
    SendMessage AgentHwnd, WM_COPYDATA, glMyHwnd, cdCopyData
End Sub



Private Function ByteArrayToString(Bytes() As Byte) As String
    Dim iUnicode As Long, i As Long, j As Long
    
    On Error Resume Next
    i = UBound(Bytes)
    
    If (i < 1) Then
        'ANSI, just convert to unicode and return
        ByteArrayToString = StrConv(Bytes, vbUnicode)
        Exit Function
    End If
    i = i + 1
    
    'Examine the first two bytes
    CopyMemory iUnicode, Bytes(0), 2
    
    If iUnicode = Bytes(0) Then 'Unicode
        'Account for terminating null
        If (i Mod 2) Then i = i - 1
        'Set up a buffer to recieve the string
        ByteArrayToString = String$(i / 2, 0)
        'Copy to string
        CopyMemory ByVal StrPtr(ByteArrayToString), Bytes(0), i
    Else 'ANSI
        ByteArrayToString = StrConv(Bytes, vbUnicode)
    End If
                    
End Function

Private Function StringToByteArray(strInput As String, _
                                Optional bReturnAsUnicode As Boolean = True, _
                                Optional bAddNullTerminator As Boolean = False) As Byte()
    
    Dim lRet As Long
    Dim bytBuffer() As Byte
    Dim lLenB As Long
    
    If bReturnAsUnicode Then
        'Number of bytes
        lLenB = LenB(strInput)
        'Resize buffer, do we want terminating null?
        If bAddNullTerminator Then
            ReDim bytBuffer(lLenB)
        Else
            ReDim bytBuffer(lLenB - 1)
        End If
        'Copy characters from string to byte array
        CopyMemory bytBuffer(0), ByVal StrPtr(strInput), lLenB
    Else
        'METHOD ONE
'        'Get rid of embedded nulls
'        strRet = StrConv(strInput, vbFromUnicode)
'        lLenB = LenB(strRet)
'        If bAddNullTerminator Then
'            ReDim bytBuffer(lLenB)
'        Else
'            ReDim bytBuffer(lLenB - 1)
'        End If
'        CopyMemory bytBuffer(0), ByVal StrPtr(strInput), lLenB
        
        'METHOD TWO
        'Num of characters
        lLenB = Len(strInput)
        If bAddNullTerminator Then
            ReDim bytBuffer(lLenB)
        Else
            ReDim bytBuffer(lLenB - 1)
        End If
        lRet = WideCharToMultiByte(CP_ACP, 0&, ByVal StrPtr(strInput), -1, ByVal VarPtr(bytBuffer(0)), lLenB, 0&, 0&)
    End If
    
    StringToByteArray = bytBuffer
    
End Function

Public Sub CollectorLedBlink(BlinkSwitch As Integer)
    
    Select Case BlinkSwitch
    Case 0 'ºñ¦â
        PostMessage glPingManagerHwnd, glMyPingMsg, MSG_PINGCOLLECTOR_BLINK_A, 0&
    Case 1 '¶À¦â
        PostMessage glPingManagerHwnd, glMyPingMsg, MSG_PINGCOLLECTOR_BLINK_B, 0&
    Case 2 '°±¤î
        PostMessage glPingManagerHwnd, glMyPingMsg, MSG_PINGCOLLECTOR_BLINK_S, 0&
    End Select
End Sub

Public Sub SayHello()
    Dim cdCopyData As COPYDATASTRUCT
    If glPingManagerHwnd = 0 Then
        Exit Sub
    End If
    cdCopyData.dwData = MSG_PINGCOLLECTOR_SAYHELLO '¦Û©w¸q¼Æ¾Ú
    cdCopyData.cbData = Len(glMyHwnd)
    cdCopyData.lpData = VarPtr(glMyHwnd)
    'MsgBox "Log agent say hello!"
    SendMessage glPingManagerHwnd, WM_COPYDATA, glMyHwnd, cdCopyData
End Sub

'Public Sub ReportPingResult()
'    Dim cdCopyData As COPYDATASTRUCT
'    Dim LenOfData As Long
'    'Dim SN As Long
'    'result, received, lost, acc rtt, cycle interval, alert count, stat. count, stat. fail count
'    'aryReportData(1 To 8, MaxNodeIndex)
'
'    cdCopyData.dwData = MSG_PINGCOLLECTOR_REPORT_PING_RESULT '¦Û©w¸q¼Æ¾Ú
'    cdCopyData.cbData = Len(aryReportData(1, 0)) * NumOfPingNode * 8 '¶Ç°eªø«×
'    cdCopyData.lpData = VarPtr(aryReportData(1, 0))
'    SendMessage glPingManagerHwnd, WM_COPYDATA, glMyHwnd, cdCopyData
'
'End Sub
Public Sub ReportPingResult()
    Dim cdCopyData As COPYDATASTRUCT
    Dim LenOfData As Long
    'result, accreceived, acclost, accrtt, alertcount, interval, failcount, ping cycle count
    'aryPingStatData(1 To 8, MaxNodeIndex)

    cdCopyData.dwData = MSG_PINGCOLLECTOR_REPORT_PING_RESULT '¦Û©w¸q¼Æ¾Ú
    cdCopyData.cbData = Len(aryPingStatData(1, 0)) * NumOfPingNode * 8 '¶Ç°eªø«×
    cdCopyData.lpData = VarPtr(aryPingStatData(1, 0))
    SendMessage glPingManagerHwnd, WM_COPYDATA, glMyHwnd, cdCopyData

End Sub
Public Sub SendEventToManager()
    Dim cdCopyData As COPYDATASTRUCT
    Dim LenOfData As Long
    Dim EventSN As Long

    On Error GoTo ErrHandler
    If NumOfPingNode = 0 Then Exit Sub
    If glPingManagerHwnd = 0 Then Exit Sub

    EventSN = GetTickCount

    If EventBufCount = 0 Then
        cdCopyData.dwData = MSG_PINGCOLLECTOR_SENDEVENT_NOEVENT '¦Û©w¸q¼Æ¾Ú
        cdCopyData.cbData = Len(EventSN) '¶Ç°eªø«×
        cdCopyData.lpData = VarPtr(EventSN)
        SendMessage glPingManagerHwnd, WM_COPYDATA, glMyHwnd, cdCopyData
        Exit Sub
    Else
        LenOfData = Len(aryEventBuf1(1, 1)) * EventBufCount * 2 '¤Gºû°}¦C
        cdCopyData.dwData = MSG_PINGCOLLECTOR_SENDEVENT_BUF1 '¦Û©w¸q¼Æ¾Ú
        cdCopyData.cbData = LenOfData '¶Ç°eªø«×
        cdCopyData.lpData = VarPtr(aryEventBuf1(1, 1)) '¤Gºû°}¦C 'vb6 column order
        SendMessage glPingManagerHwnd, WM_COPYDATA, EventSN, cdCopyData

        LenOfData = Len(aryEventBuf2(1)) * EventBufCount
        cdCopyData.dwData = MSG_PINGCOLLECTOR_SENDEVENT_BUF2 '¦Û©w¸q¼Æ¾Ú
        cdCopyData.cbData = LenOfData '¶Ç°eªø«×
        cdCopyData.lpData = VarPtr(aryEventBuf2(1))
        SendMessage glPingManagerHwnd, WM_COPYDATA, EventSN, cdCopyData
        Call ResetEventBuf
    End If
    Exit Sub
ErrHandler:
    frmMain.ShowStatus "Send Event Data Error!"
End Sub

