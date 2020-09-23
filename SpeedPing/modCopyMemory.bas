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

Public Const WM_APP = &H8000
Public Const WM_MYTESTMSG = WM_APP + 1
'Copies a block of memory from one location to another.
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReplyMessage Lib "user32" (ByVal lReply As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

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
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
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
    Dim cdCopyData As COPYDATASTRUCT
    Dim LenOfCopyData As Long
    Dim SN As Long
    Dim AgentHwnd As Long
    Dim AgentID As Long
    Dim i As Integer

    Select Case Lmsg
    Case glMyPingMsg
        Select Case wParam
        Case MSG_AGENT_BLINK_A
            frmMain.AgentLedBlink lParam, MSG_AGENT_BLINK_A
        Case MSG_AGENT_BLINK_B
            frmMain.AgentLedBlink lParam, MSG_AGENT_BLINK_B
        Case MSG_AGENT_BLINK_S
            frmMain.AgentLedBlink lParam, MSG_AGENT_BLINK_S
        Case MSG_PINGCOLLECTOR_BLINK_A
            frmMain.CollectorLedBlink MSG_PINGCOLLECTOR_BLINK_A
        Case MSG_PINGCOLLECTOR_BLINK_B
            frmMain.CollectorLedBlink MSG_PINGCOLLECTOR_BLINK_B
        Case MSG_PINGCOLLECTOR_BLINK_S
            frmMain.CollectorLedBlink MSG_PINGCOLLECTOR_BLINK_S
            
        Case MSG_PINGCOLLECTOR_SAYGOODBYE
            frmMain.PingCollectorSayGoodbye
        Case MSG_AGENT_SAYGOODBYE 'Say Goodbye®É¥Îpostmessage,Â÷¶}ªº°®¯Ü¤@ÂI
            AgentID = lParam
            frmMain.AgentSayGoodbye lParam
        End Select
    Case WM_COPYDATA
        
        Call CopyMemory(cdCopyData, ByVal lParam, Len(cdCopyData))
        'MsgBox cdCopyData.dwData
        
        LenOfCopyData = cdCopyData.cbData
        Select Case cdCopyData.dwData
            
            Case MSG_PINGCOLLECTOR_REPORT_PING_RESULT
                Call CopyMemory(aryReportData(1, 0), ByVal cdCopyData.lpData, LenOfCopyData)
                ReplyMessage 1
                frmMain.RefreshPingStatus
                
            '******************************************************************************
            Case MSG_PINGCOLLECTOR_SENDEVENT_NOEVENT
                LastLogTick = GetTickCount
                ReplyMessage 1
                'MsgBox "No Event!"
            Case MSG_PINGCOLLECTOR_SENDEVENT_BUF1
                EventSN = wParam
                Call CopyMemory(aryEventBuf1(1, 1), ByVal cdCopyData.lpData, LenOfCopyData)
                ReplyMessage 1
            Case MSG_PINGCOLLECTOR_SENDEVENT_BUF2
                If wParam <> EventSN Then
                    ReplyMessage 1
                    frmMain.StopLog
                    MsgBox "Event SN ¿ù»~! log°±¤î!"
                    Exit Function
                End If
                Call CopyMemory(aryEventBuf2(1), ByVal cdCopyData.lpData, LenOfCopyData)
                '­pºâEventªº¶µ¥Ø
                EventBufCount = LenOfCopyData / Len(aryEventBuf2(1))
                LastLogTick = GetTickCount
                ReplyMessage 1
                frmMain.SaveEvent
                
            '******************************************************************************
            
            Case MSG_PINGCOLLECTOR_READY
                ReplyMessage 1
                frmMain.PingCollectorReady
                
            Case MSG_PINGCOLLECTOR_LOADPINGLIST_OK
                ReplyMessage 1
                frmMain.PingCollectorLoadPingListOK
                
            Case MSG_PINGCOLLECTOR_LOADPINGLIST_ERR
                ReplyMessage 1
                frmMain.PingCollectorLoadPingListErr
                MsgBox "PingCollector ¸ü¤Jªº¸`ÂI¼Æ¶q¤£²Å!", vbExclamation, MsgTitle

            Case MSG_PINGCOLLECTOR_SAYHELLO
                Call CopyMemory(glPingCollectorHwnd, ByVal cdCopyData.lpData, LenOfCopyData)
                ReplyMessage 1
                frmMain.PingCollectorSayHello
                
            Case MSG_PLS_PINGCOLLECTOR_CREATEAGENT '»Ý­n¶Çoptions(ini)
                ReplyMessage 1
                
            Case MSG_AGENT_SAYHELLO
                Call CopyMemory(AgentHwnd, ByVal cdCopyData.lpData, LenOfCopyData)
                ReplyMessage 1
                'wParam = agentid
                AgentID = wParam
                aryAgentHwnd(AgentID) = AgentHwnd
                frmMain.AgentSayHello AgentID '§ïÅÜ¿O¸¹¬°¶À¦â
                
            Case MSG_AGENT_READY
                ReplyMessage 1
                AgentID = wParam
                frmMain.AgentReady AgentID
            
        End Select
'        'Call InterProcessComms(lParam)
    
    End Select

    WindowProc = CallWindowProc(LocalPrevWndProc, Lwnd, Lmsg, wParam, lParam)
End Function

Public Sub TellCollectorToReportPingResult()
    PostMessage glPingCollectorHwnd, glMyPingMsg, MSG_PLS_PINGCOLLECTOR_REPORT_PING_RESULT, 0
End Sub

Public Sub TellAppToDoSomething(cmd As Long, AgentHwnd As Long)
    Dim cdCopyData As COPYDATASTRUCT
    Dim DoCmd As Long
    DoCmd = cmd
    cdCopyData.dwData = DoCmd '¦Û©w¸q¼Æ¾Ú
    
    cdCopyData.cbData = Len(DoCmd)
    cdCopyData.lpData = VarPtr(DoCmd)
    SendMessage AgentHwnd, WM_COPYDATA, glMyHwnd, cdCopyData
End Sub
Public Sub TellAgentToClose()
    Dim i As Integer
    Dim AgentHwnd As Long
    On Error Resume Next
    If IsArrayInitialized(aryAgentHwnd) Then
        For i = LBound(aryAgentHwnd) To UBound(aryAgentHwnd)
            AgentHwnd = aryAgentHwnd(i)
            If AgentHwnd <> 0 Then
                PostMessage AgentHwnd, glMyPingMsg, MSG_AGENT_CLOSE, 0&
            End If
        Next
    End If
End Sub
Public Sub TellCollectorToClose()
    On Error Resume Next
    If glPingCollectorHwnd <> 0 Then
        PostMessage glPingCollectorHwnd, glMyPingMsg, MSG_PLS_PINGCOLLECTOR_CLOSE, 0&
    End If
End Sub
Public Sub TellCollectorSendEvent()
    On Error Resume Next
    If glPingCollectorHwnd <> 0 Then
        PostMessage glPingCollectorHwnd, glMyPingMsg, MSG_PLS_PINGCOLLECTOR_SEND_EVENT, 0&
    End If
End Sub
Public Sub TellAgentToLoadINI(AgentID As Long)
    Dim cdCopyData As COPYDATASTRUCT
    If AgentID = 0 Then
        Exit Sub
    End If
    If aryAgentHwnd(AgentID) = 0 Then
        Exit Sub
    End If
    cdCopyData.dwData = MSG_AGENT_LOAD_INI '¦Û©w¸q¼Æ¾Ú
    cdCopyData.cbData = Len(aryINI(1)) * 12
    cdCopyData.lpData = VarPtr(aryINI(1))
    SendMessage aryAgentHwnd(AgentID), WM_COPYDATA, glMyHwnd, cdCopyData
End Sub
Public Sub TellCollectorToLoadList(num As Long)
    '¶Ç°e¸ü¤Jªº¼Æ¶q¥H°µ¬°¤ñ¹ï¥Î
    Dim cdCopyData As COPYDATASTRUCT
    
    cdCopyData.dwData = MSG_PLS_PINGCOLLECTOR_LOADPINGLIST '¦Û©w¸q¼Æ¾Ú
    
    cdCopyData.cbData = Len(num)
    cdCopyData.lpData = VarPtr(num)
    SendMessage glPingCollectorHwnd, WM_COPYDATA, glMyHwnd, cdCopyData
End Sub
Public Sub TellCollectorToLoadINI()
    Dim cdCopyData As COPYDATASTRUCT
    If glPingCollectorHwnd = 0 Then Exit Sub
    cdCopyData.dwData = MSG_PLS_PINGCOLLECTOR_LOADINI '¦Û©w¸q¼Æ¾Ú
    
    cdCopyData.cbData = Len(aryINI(1)) * 12
    cdCopyData.lpData = VarPtr(aryINI(1))
    SendMessage glPingCollectorHwnd, WM_COPYDATA, glMyHwnd, cdCopyData
End Sub


