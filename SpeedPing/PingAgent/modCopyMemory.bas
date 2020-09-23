Attribute VB_Name = "modCopyMemory"
Option Explicit

Private Type COPYDATASTRUCT
    dwData As Long
    cbData As Long
    lpData As Long
End Type
Private Const MSG_MYPING_BASE As Long = 814692
Private Const MSG_AGENT_LOADPINGLIST As Long = MSG_MYPING_BASE + 1
Private Const MSG_AGENT_SAYHELLO As Long = MSG_MYPING_BASE + 2
Private Const MSG_AGENT_STARTPING As Long = MSG_MYPING_BASE + 3
Private Const MSG_AGENT_STOPPING As Long = MSG_MYPING_BASE + 4
Private Const MSG_REPORT_PINGSTATUS_RTT As Long = MSG_MYPING_BASE + 5
Private Const MSG_REPORT_PINGSTATUS_RECV As Long = MSG_MYPING_BASE + 6
Private Const MSG_REPORT_PINGSTATUS_LOST As Long = MSG_MYPING_BASE + 7
Private Const MSG_REPORT_PINGSTATUS_OK As Long = MSG_MYPING_BASE + 8
Private Const MSG_REPORT_PINGSTATUS_STATUS As Long = MSG_MYPING_BASE + 9
Private Const MSG_AGENT_READY As Long = MSG_MYPING_BASE + 10
Private Const MSG_AGENT_CLOSE As Long = MSG_MYPING_BASE + 11 '¥Ñserverµo¥X
Private Const MSG_REPORT_PINGSTATUS_INFO As Long = MSG_MYPING_BASE + 12
Private Const MSG_AGENT_BLINK As Long = MSG_MYPING_BASE + 13
Private Const MSG_AGENT_SAYGOODBYE As Long = MSG_MYPING_BASE + 14
Private Const MSG_AGENT_SET_CYCLEINTERVAL As Long = MSG_MYPING_BASE + 15
Private Const MSG_AGENT_LOAD_INI As Long = MSG_MYPING_BASE + 16
Private Const MSG_AGENT_REPORT_LOAD_INI_OK As Long = MSG_MYPING_BASE + 17
Private Const MSG_PLS_SAYHELLO As Long = MSG_MYPING_BASE + 18
Private Const MSG_REPORT_PINGSTAT As Long = MSG_MYPING_BASE + 19
Private Const MSG_AGENT_KEEPALIVE As Long = MSG_MYPING_BASE + 20
Private Const MSG_AGENT_BLINK_A As Long = MSG_MYPING_BASE + 21
Private Const MSG_AGENT_BLINK_B As Long = MSG_MYPING_BASE + 22
Private Const MSG_AGENT_BLINK_S As Long = MSG_MYPING_BASE + 23

Private Const MSG_KEEPALIVE_REQUEST As Long = MSG_MYPING_BASE + 400
Private Const MSG_PINGAGENT_KEEPALIVE_REPLY As Long = MSG_MYPING_BASE + 401
Private Const MSG_PINGCOLLECTOR_KEEPALIVE_REPLY As Long = MSG_MYPING_BASE + 402
Private Const MSG_STOP_PING As Long = MSG_MYPING_BASE + 403

Public Const MSG_MYPING_POSTMSG As String = "MsgMyPing1.0"

Private Const WM_COPYDATA = &H4A
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReplyMessage Lib "user32" (ByVal lReply As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


'//subclassing
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Const GWL_WNDPROC = (-4)
Dim LocalPrevWndProc As Long

Public Sub Hook()
    On Error Resume Next
    LocalPrevWndProc = SetWindowLong(glMyHwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub
Public Sub UnHook()
    Dim WorkFlag As Long

    On Error Resume Next
    If LocalPrevWndProc <> 0 Then
        WorkFlag = SetWindowLong(glMyHwnd, GWL_WNDPROC, LocalPrevWndProc)
    End If
End Sub

Private Function WindowProc(ByVal Lwnd As Long, ByVal Lmsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'        lParam: pointer to structure with data
'        wParam: handle of sending window
    Dim cdCopyData As COPYDATASTRUCT
    Dim LenOfCopyData As Long
    Dim i As Long
    Dim aryINI() As Long
    Dim ininum As Long
    
    Select Case Lmsg
    Case glMyPingMsg
        Select Case wParam
        Case MSG_AGENT_CLOSE
                frmMain.CloseMe
        End Select
        
    Case WM_COPYDATA
        CopyMemory cdCopyData, ByVal lParam, Len(cdCopyData)
        Select Case cdCopyData.dwData
            Case MSG_AGENT_LOADPINGLIST
                'ª`·NLoad Ping List®É¦]¬°­nredim round trip time array®É·|¥Î¨ìglPingCount,©Ò¥H­n¥ýLoad INI¤~¦æ
                LenOfCopyData = cdCopyData.cbData
                NumOfPingNode = LenOfCopyData \ Len(LenOfCopyData) '¦P¼Ë¬Olong
                If NumOfPingNode > 0 Then
                    MaxNodeIndex = NumOfPingNode - 1
                    ReDim aryInetAddr(MaxNodeIndex)
                    Call CopyMemory(aryInetAddr(0), ByVal cdCopyData.lpData, LenOfCopyData)
                    ReplyMessage 1
'                    For i = 0 To MaxNodeIndex
'                         aryIPAddress(i) = GetInetStrFromPtr(aryInetAddr(i))
'                    Next
                    frmMain.LoadPingList '-->·|°õ¦æReportReady
                Else
                    ReplyMessage 1
                End If
            Case MSG_AGENT_LOAD_INI
                LenOfCopyData = cdCopyData.cbData
                ininum = LenOfCopyData \ Len(LenOfCopyData) '¦P¼Ë¬Olong
                ReDim aryINI(1 To ininum)
                Call CopyMemory(aryINI(1), ByVal cdCopyData.lpData, LenOfCopyData)
                ReplyMessage 1

                glPingCount = aryINI(1)
                glPingTimeOutHost = aryINI(2)
                glPingMaxBurst = aryINI(3)
                glPingTimeOutBatch = aryINI(4)
                glWaitForSingleObject = aryINI(5)
                glMinCycleInterval = aryINI(6)
                glThreshold = aryINI(8)
                
                If glPingCount > 0 And _
                    glPingTimeOutHost > 0 And _
                    glPingMaxBurst > 0 And _
                    glPingTimeOutBatch > 0 And _
                    glWaitForSingleObject > 0 And _
                    glMinCycleInterval > 0 And _
                    glThreshold > 0 Then
                    
                    ReportLoadIniOK
                End If

            Case MSG_AGENT_STARTPING
                ReplyMessage 1
                frmMain.DoPing

            Case MSG_STOP_PING
                ReplyMessage 1
                frmMain.StopPing
            
            Case MSG_PLS_SAYHELLO
                ReplyMessage 1
                SayHello

        End Select

    End Select
    WindowProc = CallWindowProc(LocalPrevWndProc, Lwnd, Lmsg, wParam, lParam)
End Function
'Public Sub KeepAliveReply(SenderHwnd As Long)
'    Dim cdCopyData As COPYDATASTRUCT
'    If SenderHwnd = 0 Then
'        Exit Sub
'    End If
'    cdCopyData.dwData = MSG_PINGAGENT_KEEPALIVE_REPLY '¦Û©w¸q¼Æ¾Ú
'    cdCopyData.cbData = Len(MyAgentID)
'    cdCopyData.lpData = VarPtr(MyAgentID)
'    SendMessage SenderHwnd, WM_COPYDATA, glMyHwnd, cdCopyData
'End Sub
'Public Sub KeepAliveReply(SenderHwnd As Long)
'    If SenderHwnd = 0 Then
'        Exit Sub
'    End If
'    PostMessage SenderHwnd, glMyPingMsg, MSG_PINGAGENT_KEEPALIVE_REPLY, MyAgentID
'End Sub
Public Sub SayHello()
    Dim cdCopyData As COPYDATASTRUCT
    
    If glPingManagerHwnd = 0 Then
        Exit Sub
    End If
    
    cdCopyData.dwData = MSG_AGENT_SAYHELLO '¦Û©w¸q¼Æ¾Ú
    cdCopyData.cbData = Len(glMyHwnd)
    cdCopyData.lpData = VarPtr(glMyHwnd)
  
    SendMessage glPingManagerHwnd, WM_COPYDATA, MyAgentID, cdCopyData
    SendMessage glPingCollectorHwnd, WM_COPYDATA, MyAgentID, cdCopyData
End Sub
Public Sub ReportReady()
    Dim cdCopyData As COPYDATASTRUCT
    Dim DoCmd As Long
    DoCmd = MSG_AGENT_READY
    cdCopyData.dwData = DoCmd '¦Û©w¸q¼Æ¾Ú
    cdCopyData.cbData = Len(DoCmd)
    cdCopyData.lpData = VarPtr(DoCmd)
    SendMessage glPingManagerHwnd, WM_COPYDATA, MyAgentID, cdCopyData
    SendMessage glPingCollectorHwnd, WM_COPYDATA, MyAgentID, cdCopyData
    
    frmMain.StartBlink
End Sub
Public Sub ReportLoadIniOK()
    Dim cdCopyData As COPYDATASTRUCT
    Dim DoCmd As Long
    DoCmd = MSG_AGENT_REPORT_LOAD_INI_OK
    cdCopyData.dwData = DoCmd '¦Û©w¸q¼Æ¾Ú
    cdCopyData.cbData = Len(DoCmd)
    cdCopyData.lpData = VarPtr(DoCmd)
    '¥u»Ý¦VCollector¦^³ø --> Collector·|§i¶D§Ú --> MSG_AGENT_LOADPINGLIST
    SendMessage glPingCollectorHwnd, WM_COPYDATA, MyAgentID, cdCopyData
End Sub

Public Sub AgentLedBlink(BlinkSwitch As Integer)
    '¥u¦VManager Blink,§K±o¼vÅTCollectorªºperformance
    Select Case BlinkSwitch
    Case 0 'ºñ¦â
        PostMessage glPingManagerHwnd, glMyPingMsg, MSG_AGENT_BLINK_A, MyAgentID
        If glDebugMode Then
            PostMessage glPingCollectorHwnd, glMyPingMsg, MSG_AGENT_BLINK_A, MyAgentID
        End If
    Case 1 '¶À¦â
        PostMessage glPingManagerHwnd, glMyPingMsg, MSG_AGENT_BLINK_B, MyAgentID
        If glDebugMode Then
            PostMessage glPingCollectorHwnd, glMyPingMsg, MSG_AGENT_BLINK_B, MyAgentID
        End If
    Case 2 '°±¤î
        PostMessage glPingManagerHwnd, glMyPingMsg, MSG_AGENT_BLINK_S, MyAgentID
        If glDebugMode Then
            PostMessage glPingCollectorHwnd, glMyPingMsg, MSG_AGENT_BLINK_S, MyAgentID
        End If
    End Select
End Sub
Public Sub SayGoodbye()
    If glPingManagerHwnd <> 0 Then
        PostMessage glPingManagerHwnd, glMyPingMsg, MSG_AGENT_SAYGOODBYE, MyAgentID
    End If
End Sub
Public Sub ReportPingStatus(StartNode As Long, EndNode As Long)
    Dim cdCopyData As COPYDATASTRUCT
    Dim PingInfo(1 To 5) As Long
    Dim Num As Long
    Dim LenOfData As Long
    Dim SN As Long

    If glPingCollectorHwnd = 0 Then
        frmMain.CloseMe
        Exit Sub
    End If
    Num = EndNode - StartNode + 1
    
    SN = GetTickCount
    'MsgBox CycleInterval
    cdCopyData.dwData = MSG_REPORT_PINGSTATUS_INFO '¦Û©w¸q¼Æ¾Ú
    PingInfo(1) = glPingCount
    PingInfo(2) = StartNode
    PingInfo(3) = EndNode
    PingInfo(4) = SN
    PingInfo(5) = CycleInterval
    
    cdCopyData.cbData = Len(PingInfo(1)) * 5 '¶Ç°eªø«×
    cdCopyData.lpData = VarPtr(PingInfo(1))
    SendMessage glPingCollectorHwnd, WM_COPYDATA, MyAgentID, cdCopyData
    'Visual Basic 6 stores two-dimensional arrays in column-major order (it places the items in a column in adjacent memory locations)
    'while Visual Basic .NET stores them in row-major order (it places the items in a row in adjacent memory locations).
    
    LenOfData = Len(LenOfData) * Num * 6 '¬Ò¬°long
    cdCopyData.dwData = MSG_REPORT_PINGSTAT '¦Û©w¸q¼Æ¾Ú
    cdCopyData.cbData = LenOfData '¶Ç°eªø«×
    cdCopyData.lpData = VarPtr(aryAgentPingResultData(1, StartNode)) '¤Gºû°}¦C 'vb6 column order
    SendMessage glPingCollectorHwnd, WM_COPYDATA, MyAgentID, cdCopyData
    
    cdCopyData.dwData = MSG_REPORT_PINGSTATUS_OK '¦Û©w¸q¼Æ¾Ú
    cdCopyData.cbData = Len(SN)
    cdCopyData.lpData = VarPtr(SN)
    SendMessage glPingCollectorHwnd, WM_COPYDATA, MyAgentID, cdCopyData
End Sub
