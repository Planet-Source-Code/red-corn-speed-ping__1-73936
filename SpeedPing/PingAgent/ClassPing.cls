VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassPing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Type ip_option_information
    ttl             As Byte 'Byte     'Time To Live
    Tos             As Byte     'Type Of Service
    flags           As Byte     'IP header flags
    OptionsSize     As Byte     'Size in bytes of options data
    OptionsData     As Long     'Pointer to options data
End Type

Private Type ICMP_ECHO_REPLY
    Address         As Long
    status          As Long
    RoundTripTime   As Long
    datasize        As Long 'formerly integer
   'Reserved        As Integer
    DataPointer     As Long
    Options         As ip_option_information
    Data            As String * 250
End Type

Private Declare Function IcmpCreateFile Lib "icmp.dll" () As Long
Private Declare Function IcmpCloseHandle Lib "icmp.dll" (ByVal IcmpHandle As Long) As Long
'Private Declare Function inet_addr Lib "wsock32" (ByVal s As String) As Long
'The IcmpSendEcho2 function is called asynchronously when either the ApcRoutine or Event parameters are specified.
Private Declare Function IcmpSendEcho2 Lib "icmp.dll" (ByVal IcmpHandle As Long, ByVal hEvent As Long, ByVal ApcRoutine As Any, ByVal ApcContext As Long, ByVal DestinationAddress As Long, ByVal RequestData As String, ByVal RequestSize As Long, ByVal RequestOptions As Long, ByVal aryReplyBuffer As Long, ByVal ReplySize As Long, ByVal TimeOut As Long) As Long
Private Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (lpEventAttributes As SECURITY_ATTRIBUTES, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As String) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Declare Function ResetEvent Lib "kernel32" (ByVal hEvent As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const INVALID_HANDLE_VALUE = -1
Private Const WAIT_OBJECT_0 As Long = 0
Private Const WAIT_TIMEOUT As Long = &H102
Private Const WAIT_ABANDONED = &H80
Private Const WAIT_FAILED = -1&
Private Const sSendData As String = "MYPING0101"


Private SecurityAttribute As SECURITY_ATTRIBUTES
Private LenSendData As Long
Private LenBuffer As Long
Private EmptyBuffer As ICMP_ECHO_REPLY


Private TickThreshold() As Long
Public Event PingFinished(NodeIndexLow As Long, NodeIndexHigh As Long)

Private hIcmp() As Long
Private hEvent() As Long
Private EchoSent() As Integer
Private Buffer() As ICMP_ECHO_REPLY
Private EchoProcessed() As Integer

Public Function PingHostList(NodeIndexLow As Long, NodeIndexHigh As Long) As Boolean
    Dim NumOfTestNodes As Long
    
    Dim NumOfProcessed As Long
    Dim i As Long
    Dim NodeIndex As Long
    
    Dim Tick1 As Long
    Dim Tick2 As Long
    Dim TickUsed As Long
    Dim TimeOut As Long
    Dim WaitRes As Long
    Dim rtt As Long
    Dim nEnd As Long
    'On Error GoTo ErrHandler
    
    CycleStart = GetTickCount
    
    NumOfTestNodes = NodeIndexHigh - NodeIndexLow + 1
    TimeOut = glPingTimeOutBatch * glPingCount
    'MsgBox TimeOut
    nEnd = NumOfTestNodes - 1
    
'//init the needed arrays
  
   
    ReDim hIcmp(nEnd)
    ReDim hEvent(nEnd)
    
    ReDim EchoSent(nEnd)
    ReDim Buffer(nEnd)
    ReDim EchoProcessed(nEnd)
    
    
    
    For i = 0 To nEnd
        '//We create an Event for WaitForMultipleObjects to wait for
        hEvent(i) = CreateEvent(SecurityAttribute, True, False, MyEventID & i)
        hIcmp(i) = IcmpCreateFile()    '//Standard for ICMP
    Next
    Tick1 = GetTickCount
    NumOfProcessed = 0  '//Reset Counter for processed hosts
    'MsgBox "nEnd = " & nEnd
    Do
        For i = 0 To nEnd
            NodeIndex = NodeIndexLow + i
            '//Did we already request an echo
            '//·í¸ÓEcho Sent,¥B¤w¸gProcessed(EchoSent+1,µM«áEchoProcessed³B²z«á+1)
            If (EchoSent(i) = EchoProcessed(i)) And (EchoSent(i) < glPingCount) Then
                EchoSent(i) = EchoSent(i) + 1    '//Lock this adress, so we don't create a second Ping
                Call IcmpSendEcho2(hIcmp(i), hEvent(i), Nothing, 0, aryInetAddr(NodeIndex), sSendData, LenSendData, 0, VarPtr(Buffer(i)), LenBuffer, glPingTimeOutHost)
            End If
            If EchoProcessed(i) < glPingCount Then
                WaitRes = WaitForSingleObject(hEvent(i), glWaitForSingleObject) '//We wait only for 5 Millisecs, to give all the other IPs Time to answer
                Select Case WaitRes
                    Case WAIT_OBJECT_0
                        DoEvents
                        EchoProcessed(i) = EchoProcessed(i) + 1
                        NumOfProcessed = NumOfProcessed + 1

                        If Buffer(i).status = 0 Then '¤w¦^À³
                            '//Valid Info : So Host is reachable
                            aryAgentPingResultData(6, NodeIndex) = MY_PING_OK 'status¤£¥ÎIP_SUCCESS(­È0),§ï¥Î¦Û©w¸qªº±`¼Æ
                            rtt = Buffer(i).RoundTripTime
                            'Min RTT
                            If rtt < aryAgentPingResultData(1, NodeIndex) Then
                                aryAgentPingResultData(1, NodeIndex) = rtt
                            End If
                            'Avg RTT
                            aryAgentPingResultData(2, NodeIndex) = aryAgentPingResultData(2, NodeIndex) + rtt '³Ì«á¦A¥­§¡
                            'Max RTT
                            If rtt > aryAgentPingResultData(3, NodeIndex) Then
                                aryAgentPingResultData(3, NodeIndex) = rtt
                            End If
                            
                            aryAgentPingResultData(4, NodeIndex) = aryAgentPingResultData(4, NodeIndex) + 1
                            'RaiseEvent PingSuccess(NodeIndex, Buffer(i).RoundTripTime)
                        Else
                            aryAgentPingResultData(6, NodeIndex) = Buffer(i).status
                            aryAgentPingResultData(5, NodeIndex) = aryAgentPingResultData(5, NodeIndex) + 1
                            'RaiseEvent PingFail(NodeIndex, Buffer(i).status)
                        End If
                        ResetEvent hEvent(i)
                        Buffer(i) = EmptyBuffer
                    Case Else
                        'Echo ÁÙ¨S¦³¦^À³,¬Ý¬Ý¬O§_¹O®É
                        DoEvents
                        Tick2 = GetTickCount
                        TickUsed = TickDiff(Tick1, Tick2)
                        If TickUsed > (TickThreshold(EchoSent(i))) Then
                            aryAgentPingResultData(6, NodeIndex) = PING_NOT_YET
                            aryAgentPingResultData(5, NodeIndex) = aryAgentPingResultData(5, NodeIndex) + 1
                            EchoProcessed(i) = EchoProcessed(i) + 1
                            NumOfProcessed = NumOfProcessed + 1
                            ResetEvent hEvent(i)
                            Buffer(i) = EmptyBuffer
                            'CopyMemory Buffer(i), ByVal EmptyBuffer, Len(EmptyBuffer)
                        End If
                End Select
            End If
        
            DoEvents
            
        Next
        Tick2 = GetTickCount
        TickUsed = TickDiff(Tick1, Tick2)
        
        If TickUsed >= TimeOut Then
            'MsgBox "timeout!!"
            DoEvents
            Exit Do
        End If
        
        '//All done, no more Action needed
        If NumOfProcessed >= (NumOfTestNodes * glPingCount) Then
            Exit Do
        End If
    Loop While (1)
        
'//IMPORTANT : UNregister all messages/Evemnts and open ICMP-handles
    For i = 0 To nEnd
        'ResetEvent hEvent(i)
        CloseHandle hEvent(i)
        IcmpCloseHandle (hIcmp(i))
        NodeIndex = NodeIndexLow + i
        If EchoProcessed(i) < glPingCount Then
            '//notify caller with the rest of the host, that we haven't hit
            aryAgentPingResultData(6, NodeIndex) = PING_NOT_YET
            aryAgentPingResultData(5, NodeIndex) = glPingCount - aryAgentPingResultData(4, NodeIndex)
        End If
        '­pºâ²Î­p­È
        '¦³µL¦¬¨ì«Ê¥]
        If aryAgentPingResultData(4, NodeIndex) > 0 Then
        '¦³¦¬¨ì
            'Avg RTT
            '§ï¥ÑPingCollector°µ²Î­p
            'aryAgentPingResultData(2, NodeIndex) = aryAgentPingResultData(2, NodeIndex) / aryAgentPingResultData(4, NodeIndex)
            '
            'ÅýServer¦¬¨ì¦¨¥\©Î¥¢±Ñªºstatus­È
            If aryAgentPingResultData(5, NodeIndex) >= glThreshold Then
                '¥i¯à³Ì«á¤@¦¸ping¬OOKªº
                If aryAgentPingResultData(6, NodeIndex) = MY_PING_OK Then
                    aryAgentPingResultData(6, NodeIndex) = IP_REQ_TIMED_OUT
                End If
            Else
                aryAgentPingResultData(6, NodeIndex) = MY_PING_OK
            End If
        Else
        '¨S¦¬¨ì
            aryAgentPingResultData(1, NodeIndex) = -1 'min
            aryAgentPingResultData(2, NodeIndex) = -1 'avg
            aryAgentPingResultData(3, NodeIndex) = -1 'max
        End If
    Next
'    MsgBox "a=" & aryAgentPingResultData(4, 0) & " ; b= " & aryAgentPingResultData(4, 1)
'    MsgBox "1=" & aryAgentPingResultData(1, 0) & "; 2=" & aryAgentPingResultData(1, 1)
    'aryAgentPingResultData(1, 0) = -1
'//And we free up some memory
''    Erase hIcmp
''    Erase hEvent
''
''    Erase EchoSent
''    Erase Buffer
''    Erase EchoProcessed
    
    CycleEnd = GetTickCount
    CycleInterval = TickDiff(CycleStart, CycleEnd) 'GetTickCount ¥i¯à¤p©ó0

    '¤£¥i¥H¦b¦¹§PÂ_If PingCount = glPingCount Then,¦]¬°tmrPing¤£·|Ä~Äò§@¥Î
    'MsgBox "pingok"
    RaiseEvent PingFinished(NodeIndexLow, NodeIndexHigh)
    
    Exit Function
'ErrHandler:
'    MsgBox "PingHostList Error!" & vbCrLf & NodeIndex & vbCrLf & i
End Function

Public Sub InitPing()
    Dim i As Integer
    ReDim TickThreshold(1 To glPingCount)
    For i = 1 To glPingCount
      TickThreshold(i) = glPingTimeOutBatch * i
    Next
End Sub
Private Sub Class_Initialize()
    '//needed for Create Event
    With SecurityAttribute
        .nLength = Len(SecurityAttribute) 'we pass the length of SecurityAttribute
        .lpSecurityDescriptor = 0
        .bInheritHandle = 0
    End With
    Dim Buffer(1) As ICMP_ECHO_REPLY
    
    LenSendData = Len(sSendData)
    LenBuffer = Len(Buffer(0))
    
End Sub
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


