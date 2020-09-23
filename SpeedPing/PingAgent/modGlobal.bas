Attribute VB_Name = "modGlobal"
Option Explicit
Public Const MsgTitle As String = "Ping Agent"

Public glMyHwnd As Long
Public MyAgentID As Long
Public MyEventID As String
Public glDebugMode As Boolean
Public glPingCollectorHwnd As Long
Public glMyPingMsg As Long

Public CycleInterval As Long
Public CycleStart As Long
Public CycleEnd As Long

Public NumOfPingNode As Long
Public MaxNodeIndex As Long
Public PingIsRunning As Boolean
Public glPingTimeOutHost As Long ' = 1000 ' ms
Public glPingTimeOutBatch As Long '= 4000 'ms
Public glPingMaxBurst As Long '= 500 '¨C­Óinterval°e¥Xªºpacket¼Æ¶q
Public glPingCount As Long '= 5 '¨C­Ó¸`ÂI´ú¸Õªº¦¸¼Æ
Public glWaitForSingleObject As Long '= 2 'ms
Public glPingManagerHwnd As Long
Public glMinCycleInterval As Long 'ms
Public glThreshold As Integer
Public UserStop As Boolean
Public UserClose As Boolean
Public NextPingIsWaiting As Boolean

Public aryInetAddr() As Long

Public aryAgentPingResultData() As Long 'minrtt, avgrtt, maxrtt, received, lost, status

Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long


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

'Public Declare Function inet_addr Lib "wsock32" (ByVal s As String) As Long

