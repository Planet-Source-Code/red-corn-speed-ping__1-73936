VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  '沒有框線
   Caption         =   "Ping"
   ClientHeight    =   900
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   2145
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   900
   ScaleWidth      =   2145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '系統預設值
   Visible         =   0   'False
   Begin VB.Timer tmrCloseMe 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1470
      Top             =   270
   End
   Begin VB.Timer tmrBlink 
      Enabled         =   0   'False
      Interval        =   1270
      Left            =   900
      Top             =   270
   End
   Begin VB.Timer tmrPing 
      Left            =   300
      Top             =   270
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents cPing As ClassPing
Attribute cPing.VB_VarHelpID = -1
Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long

Private LastPingEndNode As Long
Private PingStartNode As Long, PingEndNode As Long

Private LastCheckEnd As Long
Private BlinkSwitch As Integer
Private RndNum(0 To 9) As Long
Private BlinkInterval As Long

Public Sub DoPing()
   If NumOfPingNode = 0 Then
        Exit Sub
    End If
    If PingIsRunning Then
        Exit Sub
    End If
    PingIsRunning = True
    LastPingEndNode = MaxNodeIndex '將指標移至最後
    LastCheckEnd = MaxNodeIndex
    UserStop = False
    Call cPing.InitPing
    BlinkSwitch = 0 '開始閃爍
    tmrPing.Interval = 5
    tmrPing.Enabled = True
End Sub

Public Sub StopPing()
    On Error Resume Next
    UserStop = True
'    MsgBox "UserStop = True"
End Sub

Private Sub cPing_PingFinished(NodeIndexLow As Long, NodeIndexHigh As Long)
    
    If Not UserStop And Not UserClose Then
        ReportPingStatus NodeIndexLow, NodeIndexHigh
        DoEvents
'        MsgBox "CycleInterval=" & CycleInterval
'        MsgBox "glMinCycleInterval=" & glMinCycleInterval
        If CycleInterval < glMinCycleInterval Then
            
            If CycleInterval >= 0 Then '理論上不可能<0,但因為TickDiff為從網路copy過來,尚未驗證(2010-12-05:ping自己會等於0)
                tmrPing.Interval = glMinCycleInterval - CycleInterval
            Else
                'interval 不變
            End If
        Else
            tmrPing.Interval = 5
        End If
        NextPingIsWaiting = True
        tmrPing.Enabled = True
        Exit Sub
    Else
        PingIsRunning = False
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandler
    App.TaskVisible = False
    App.Title = ""
    glMyHwnd = Me.hwnd
    glMyPingMsg = RegisterWindowMessage(MSG_MYPING_POSTMSG)
    If glMyPingMsg = 0 Then
        Unload Me
    End If
    UserClose = False
    RndNum(0) = 1390
    RndNum(1) = 1730
    RndNum(2) = 1270
    RndNum(3) = 1570
    RndNum(4) = 1370
    RndNum(5) = 1630
    RndNum(6) = 1490
    RndNum(7) = 1510
    RndNum(8) = 1310
    RndNum(9) = 1670
    
    BlinkInterval = 1270
    '取得命令列參數
    Dim cmdline As String
    Dim aryPara() As String
    Dim filename As String
    cmdline = Trim(Command)
    If cmdline <> "" Then
        aryPara = Split(cmdline, Space(1))
        If UBound(aryPara) <> 3 Then
            MsgBox "載入Ping Agent程式時發生參數數目錯誤!", vbCritical, MsgTitle
            Unload Me
        End If
        glPingManagerHwnd = aryPara(0)
        glPingCollectorHwnd = aryPara(2)
        If aryPara(3) = 1 Then
            glDebugMode = True
        Else
            glDebugMode = False
        End If
        If IsWindow(glPingManagerHwnd) = 0 Then
            'MsgBox "找不到Ping Server!", vbCritical, MsgTitle
            CloseMe
        End If
        
        MyAgentID = aryPara(1)
        'MsgBox "MyAgentID=" & MyAgentID
        BlinkInterval = RndNum(MyAgentID Mod 10)
        
        '確保唯一
        MyEventID = "PAGENT" & CStr(glMyHwnd) 'Format(MyAgentID, "00000")
        'Load pinglist file
'        filename = App.Path & "\tmp\pinglist-" & MyAgentID & ".txt"
'        LoadPingListFile filename
        Hook

        Call SayHello
    Else
        End
    End If
    
    Set cPing = New ClassPing
    
    Exit Sub
ErrHandler:
    MsgBox "Program load error!" & vbCrLf & Err.Description, vbCritical, MsgTitle
    Unload Me
End Sub
Public Sub StartBlink()
    BlinkSwitch = 0
    tmrBlink.Interval = BlinkInterval
    tmrBlink.Enabled = True
End Sub
Public Sub LoadPingList()
    If NumOfPingNode > 0 Then
        ReDim aryAgentPingResultData(1 To 6, MaxNodeIndex)
        
''        '*****Debug
''        '要加進 modIPStuff, 才能使用GetInetStrFromPtr
''        '用到FileSystemObject
''        Dim fs As FileSystemObject
''        Dim f As TextStream
''        Dim filename As String
''        Dim foldername As String
''        Dim i As Integer
''        Dim aryIPAddress() As String
''
''        Set fs = New FileSystemObject
''
''        foldername = App.Path & "\tmp"
''        If Not fs.FolderExists(foldername) Then
''            fs.CreateFolder foldername
''        End If
''        filename = foldername & "\pinglist-" & MyAgentID & ".txt"
''        Set f = fs.CreateTextFile(filename, True)
''
''        ReDim aryIPAddress(MaxNodeIndex)
''        For i = 0 To MaxNodeIndex
''            aryIPAddress(i) = GetInetStrFromPtr(aryInetAddr(i))
''            f.WriteLine i & vbTab & aryIPAddress(i) & vbTab & aryInetAddr(i)
''        Next
''        f.Close
''
''        '*****
        
        ReportReady
  
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    tmrPing.Enabled = False
    tmrBlink.Enabled = False
    tmrCloseMe.Enabled = False
    SayGoodbye
    UnHook
End Sub

Private Sub tmrBlink_Timer()
    '99/12/20改成一Load即Blink,由接收者自行判斷要不要blink,這樣就不需要keepalive了
    If PingIsRunning Then
        If UserStop And NextPingIsWaiting Then  'UserClose時tmrBlink已經disable,故在此不用考慮
            tmrPing.Enabled = False
            PingIsRunning = False
        Else
            BlinkSwitch = IIf(BlinkSwitch = 0, 1, 0)
            AgentLedBlink BlinkSwitch
            Exit Sub
        End If
    End If
    
    If BlinkSwitch = 2 Then
        AgentLedBlink BlinkSwitch
        'MsgBox "BlinkSwitch = 2"
    Else
        If BlinkSwitch = 0 Then '有可能已經是綠色
            BlinkSwitch = 2
            AgentLedBlink BlinkSwitch
        Else
            AgentLedBlink 0 '先閃成綠色
            BlinkSwitch = 2
        End If
    End If
End Sub



Private Sub tmrPing_Timer()
    Dim i As Integer
    'MsgBox "test"
    On Error GoTo ErrHandler
    NextPingIsWaiting = False
    tmrPing.Enabled = False
    If Not UserStop And Not UserClose Then
        
        PingStartNode = IIf(LastPingEndNode = MaxNodeIndex, 0, LastPingEndNode + 1)
        PingEndNode = PingStartNode + glPingMaxBurst - 1
        If PingEndNode > MaxNodeIndex Then
            PingEndNode = MaxNodeIndex
        End If
        LastPingEndNode = PingEndNode
        If PingStartNode = 0 Then
            ReDim aryAgentPingResultData(1 To 6, MaxNodeIndex)
            For i = 0 To MaxNodeIndex
                aryAgentPingResultData(1, i) = MAX_LONG_VALUE 'Min RTT
            Next
        End If
        
        cPing.PingHostList PingStartNode, PingEndNode
    Else
        '可能在已經ping完成,但還在wait當中,因為有設定glMinCycleInterval
        PingIsRunning = False
    End If
    Exit Sub
ErrHandler:
    'do nothing
    '測試時,皆為停止時未執行完成導致錯誤
End Sub

Public Sub CloseMe()
    UserClose = True
    If NextPingIsWaiting Then
        PingIsRunning = False
    End If
    tmrBlink.Enabled = False
    tmrCloseMe.Enabled = True
End Sub
Private Sub tmrCloseMe_Timer()
    If PingIsRunning Then
       'MsgBox "PingIsRunning"
    Else
        'MsgBox "CloseMe"
        tmrCloseMe.Enabled = False
        Set cPing = Nothing
        Call SayGoodbye
        Unload Me
    End If
End Sub

