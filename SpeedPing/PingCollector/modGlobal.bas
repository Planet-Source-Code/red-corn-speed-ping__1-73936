Attribute VB_Name = "modGlobal"
Option Explicit
Public Const MsgTitle As String = "Jimmy's PingCollector 1.0 (build 2010-12-1a)"
Public Const AppName As String = "JimmyPing01"
Public Const MaxAgent As Integer = 120
Public ConnStr As String

Public LoadListIntoDBOK As Boolean
Public CheckDupNodeOK As Boolean

Public PingIsRunning As Boolean

Public glPingManagerHwnd As Long
Public glMyHwnd As Long
Public Const MAX_PING_NODES As Long = 30000 '³Ì¦h3¸U
Public NumOfPingNode As Long
Public MaxNodeIndex As Long

Public glPingTimeOutHost As Long ' = 1000 ' ms
Public glPingTimeOutBatch As Long '= 4000 'ms
'Public glPingInterval As Long ' = 1200 ' ms
Public glPingMaxBurst As Long '= 500 '¨C­Óinterval°e¥Xªºpacket¼Æ¶q
Public glPingCount As Integer '= 5 '¨C­Ó¸`ÂI´ú¸Õªº¦¸¼Æ
Public glThreshold As Integer '= 5 '±¼´X­Ó¥]¬°ping lost
'Public glCheckInterval As Long '= 1200 'ms
Public glWaitForSingleObject As Long '= 2 'ms
Public glDelayStart As Long 'ms
Public glRefreshCycle As Long 'ms
Public glCycleInterval As Long 'ms
Public glContinuedFailAsDown As Integer '¨ú®ø¤£¥Î,§ï¥ÎglStatisticsCycle(°Ñ¼Æ·Ó¶Ç¹L¨Ó¨SÃö«Y)
Public glStatisticsCycle As Long

Public glLogAgentHwnd As Long
Public glDebugMode As Boolean

Public glPingListFile As String
'//Ãö©óagent
Public glAgentCount As Integer 'Ping agent's ­Ó¼Æ
Public aryAgentLB() As Long
Public aryAgentUB() As Long
Public aryAgentNumOfPingNode() As Long
Public aryAgentHwnd() As Long
Public aryAgentPingInfo() As Long
Public arySN() As Long '¶×¥Xagent¸ê®Æ®É¥Î
'Public aryAgentBlinkTick() As Long
Public aryAgentRefreshRange() As Long
Public aryAgentReady() As Boolean


Public aryINI(1 To 12) As Long
Public LogBuf As String

'//********************************************

Public aryNodeName() As String
Public aryDescription() As String
Public aryIPAddress() As String
Public aryInetAddr() As Long

Public aryAgentPingResultData() As Long

'²Î­p­È,¥Î¨Ó°Olog¥Î
Public aryNodeLastStatisticsCycleTick() As Long

Public Enum PingResult
    RESULT_UNKNOWN = 0
    RESULT_SUCCESS = 1
    RESULT_WARN = 2
    RESULT_DOWN = 3
End Enum
'¥Î¨Ó§PÂ_¤W¤@¦¸¬ODown / Up, ¦pªG¥ÑUp --> Down ©Î Down --> Up,«h°Olog
Public aryLastUpDown() As PingResult
'¥Î¨Ó°O¿ý¤W¤@¦¸ªºresult,¥Î¦bÅã¥Ü¿O¸¹®É,¦pªG¥Ø«eªºresult©M¤W¤@­Ó¿O¸¹¬Û¦P«h¤£¥Î­«½Æ³]©w,¥H­P©ó¼vÅTÅã¥Ü³t«×
Public aryLastLedResult() As PingResult

''Public aryResult() As PingResult
''Public aryAccReceived() As Integer
''Public aryAccLost() As Integer
''Public aryFailCount() As Long
''Public arySuccessCount() As Long
''Public aryContinuedFail() As Long
''Public aryDownCount() As Long
''Public aryUpDownCount() As Long
Public aryContinuedFail() As Long
'result, accreceived, acclost, accrtt, alertcount, interval, failcount, ping cycle count
Public aryPingStatData() As Long
'result, sent, received, lost, acc rtt, cycle interval, alert count
'Public aryReportData() As Long '³o¬O­n¶Çµ¹managerªº

Public glMyPingMsg As Long

Public Const MAX_EVENT_BUF As Integer = 5000
Public aryEventBuf1(1 To 2, MAX_EVENT_BUF) As Integer 'SN, EventType
Public aryEventBuf2(1 To MAX_EVENT_BUF) As Date
Public EventBufCount As Integer

Public Const MAX_LONG_VALUE As Long = 2147483647

Public Const IP_SUCCESS As Long = 0
Public Const IP_STATUS_BASE As Long = 11000
Public Const IP_BUF_TOO_SMALL As Long = (11000 + 1)
Public Const IP_DEST_NET_UNREACHABLE As Long = (11000 + 2)
Public Const IP_DEST_HOST_UNREACHABLE As Long = (11000 + 3)
Public Const IP_DEST_PROT_UNREACHABLE As Long = (11000 + 4)
Public Const IP_DEST_PORT_UNREACHABLE As Long = (11000 + 5)
Public Const IP_NO_RESOURCES As Long = (11000 + 6)
Public Const IP_BAD_OPTION As Long = (11000 + 7)
Public Const IP_HW_ERROR As Long = (11000 + 8)
Public Const IP_PACKET_TOO_BIG As Long = (11000 + 9)
Public Const IP_REQ_TIMED_OUT As Long = (11000 + 10)
Public Const IP_BAD_REQ As Long = (11000 + 11)
Public Const IP_BAD_ROUTE As Long = (11000 + 12)
Public Const IP_TTL_EXPIRED_TRANSIT As Long = (11000 + 13)
Public Const IP_TTL_EXPIRED_REASSEM As Long = (11000 + 14)
Public Const IP_PARAM_PROBLEM As Long = (11000 + 15)
Public Const IP_SOURCE_QUENCH As Long = (11000 + 16)
Public Const IP_OPTION_TOO_BIG As Long = (11000 + 17)
Public Const IP_BAD_DESTINATION As Long = (11000 + 18)
Public Const IP_ADDR_DELETED As Long = (11000 + 19)
Public Const IP_SPEC_MTU_CHANGE As Long = (11000 + 20)
Public Const IP_MTU_CHANGE As Long = (11000 + 21)
Public Const IP_UNLOAD As Long = (11000 + 22)
Public Const IP_ADDR_ADDED As Long = (11000 + 23)
Public Const IP_GENERAL_FAILURE As Long = (11000 + 50)
Public Const MAX_IP_STATUS As Long = (11000 + 50)
Public Const IP_PENDING As Long = (11000 + 255)
Public Const PING_TIMEOUT As Long = 500
Public Const INADDR_NONE As Long = &HFFFFFFFF '-1

Public Const PING_NOT_YET As Long = -101 '¦Û©w¸qstatus
Public Const MY_PING_OK As Long = 1 '¦Û©w¸qstatus

Public Declare Function inet_addr Lib "wsock32" (ByVal s As String) As Long
Public Function GetStatusCode(status As Long) As String

   Dim msg As String
   
   Select Case status
      Case IP_SUCCESS:               msg = "ip success"
      Case INADDR_NONE:              msg = "inet_addr: bad IP format"
      Case IP_BUF_TOO_SMALL:         msg = "ip buf too_small"
      Case IP_DEST_NET_UNREACHABLE:  msg = "ip dest net unreachable"
      Case IP_DEST_HOST_UNREACHABLE: msg = "ip dest host unreachable"
      Case IP_DEST_PROT_UNREACHABLE: msg = "ip dest prot unreachable"
      Case IP_DEST_PORT_UNREACHABLE: msg = "ip dest port unreachable"
      Case IP_NO_RESOURCES:          msg = "ip no resources"
      Case IP_BAD_OPTION:            msg = "ip bad option"
      Case IP_HW_ERROR:              msg = "ip hw_error"
      Case IP_PACKET_TOO_BIG:        msg = "ip packet too_big"
      Case IP_REQ_TIMED_OUT:         msg = "ip req timed out"
      Case IP_BAD_REQ:               msg = "ip bad req"
      Case IP_BAD_ROUTE:             msg = "ip bad route"
      Case IP_TTL_EXPIRED_TRANSIT:   msg = "ip ttl expired transit"
      Case IP_TTL_EXPIRED_REASSEM:   msg = "ip ttl expired reassem"
      Case IP_PARAM_PROBLEM:         msg = "ip param_problem"
      Case IP_SOURCE_QUENCH:         msg = "ip source quench"
      Case IP_OPTION_TOO_BIG:        msg = "ip option too_big"
      Case IP_BAD_DESTINATION:       msg = "ip bad destination"
      Case IP_ADDR_DELETED:          msg = "ip addr deleted"
      Case IP_SPEC_MTU_CHANGE:       msg = "ip spec mtu change"
      Case IP_MTU_CHANGE:            msg = "ip mtu_change"
      Case IP_UNLOAD:                msg = "ip unload"
      Case IP_ADDR_ADDED:            msg = "ip addr added"
      Case IP_GENERAL_FAILURE:       msg = "ip general failure"
      Case IP_PENDING:               msg = "ip pending"
      Case PING_TIMEOUT:             msg = "ping timeout"
      Case PING_NOT_YET:             msg = "ping not yet"
      Case Else:                     msg = "unknown  msg returned"
   End Select
   
   'GetStatusCode = CStr(status) & "   [ " & msg & " ]"
   GetStatusCode = msg
End Function


Public Function FindAgentID(AgentHwnd As Long) As Long
    Dim i As Long
    i = 0
    For i = 1 To glAgentCount
        'MsgBox aryAgentHwnd(i)
        If aryAgentHwnd(i) = AgentHwnd Then
            FindAgentID = i
            Exit For
        End If
    Next
End Function
Public Function TickDiff( _
    ByVal TickStart As Currency, _
    ByVal TickEnd As Currency) As Long

    ' CCur(2 ^ 32)
    Const TwoToThe32nd As Currency = 4294967296@

    ' Handle two's complement for values larger than
    ' 2147483647&
    If TickStart < 0 Then
        TickStart = TickStart + TwoToThe32nd
    End If
    ' Handle two's complement AND the case where
    ' timeGetTime/GetTickCount wraps at (2 ^ 32)ms,
    ' or ~49.7 days:
    If (TickEnd < 0) Or (TickEnd < TickStart) Then
        TickEnd = TickEnd + TwoToThe32nd
    End If
    ' Return the result
    TickDiff = TickEnd - TickStart
End Function

Public Sub ResetEventBuf()
    Erase aryEventBuf1
    Erase aryEventBuf2
    EventBufCount = 0
End Sub
