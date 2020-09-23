VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Âù½u©T©w¹ï¸Ü¤è¶ô
   Caption         =   "¿ï¶µ"
   ClientHeight    =   8940
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   5985
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '¿Ã¹õ¤¤¥¡
   Begin TabDlg.SSTab sstabOptions 
      Height          =   8145
      Left            =   105
      TabIndex        =   3
      Top             =   120
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   14367
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "ping parameters"
      TabPicture(0)   =   "frmOptions.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "picOptions(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.PictureBox picOptions 
         BorderStyle     =   0  '¨S¦³®Ø½u
         Height          =   7650
         Index           =   0
         Left            =   120
         ScaleHeight     =   7650
         ScaleWidth      =   5460
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   405
         Width           =   5460
         Begin VB.Frame fra1 
            Height          =   7530
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   15
            Width           =   5235
            Begin VB.CheckBox chkDebugMode 
               Caption         =   "run in debug mode"
               Height          =   210
               Left            =   225
               TabIndex        =   29
               Top             =   4935
               Width           =   3525
            End
            Begin VB.TextBox txtStatisticsCycle 
               Alignment       =   1  '¾a¥k¹ï»ô
               Height          =   270
               Left            =   240
               TabIndex        =   28
               Text            =   "5"
               Top             =   5910
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.TextBox txtContinuedFailAsDown 
               Alignment       =   1  '¾a¥k¹ï»ô
               Height          =   270
               Left            =   225
               TabIndex        =   26
               Text            =   "5"
               Top             =   4515
               Width           =   615
            End
            Begin VB.TextBox txtThreshold 
               Alignment       =   1  '¾a¥k¹ï»ô
               Height          =   270
               Left            =   225
               TabIndex        =   15
               Text            =   "5"
               Top             =   2145
               Width           =   615
            End
            Begin VB.TextBox txtPingTimeOutHost 
               Alignment       =   1  '¾a¥k¹ï»ô
               Height          =   270
               Left            =   225
               TabIndex        =   14
               Text            =   "1000"
               Top             =   2730
               Width           =   615
            End
            Begin VB.TextBox txtPingCount 
               Alignment       =   1  '¾a¥k¹ï»ô
               Height          =   270
               Left            =   225
               TabIndex        =   13
               Text            =   "5"
               Top             =   1590
               Width           =   615
            End
            Begin VB.TextBox txtPingMaxBurst 
               Alignment       =   1  '¾a¥k¹ï»ô
               Height          =   270
               Left            =   360
               TabIndex        =   12
               Text            =   "220"
               Top             =   6510
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.TextBox txtPingTimeOutBatch 
               Alignment       =   1  '¾a¥k¹ï»ô
               Height          =   270
               Left            =   225
               TabIndex        =   11
               Text            =   "6000"
               Top             =   3315
               Width           =   615
            End
            Begin VB.TextBox txtWaitForSingleObject 
               Alignment       =   1  '¾a¥k¹ï»ô
               Height          =   270
               Left            =   165
               TabIndex        =   10
               Text            =   "5"
               Top             =   6930
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.TextBox txtAgentCount 
               Alignment       =   1  '¾a¥k¹ï»ô
               Height          =   270
               Left            =   225
               TabIndex        =   9
               Text            =   "2"
               Top             =   1018
               Width           =   615
            End
            Begin VB.TextBox txtDelayStart 
               Alignment       =   1  '¾a¥k¹ï»ô
               Height          =   270
               Left            =   915
               TabIndex        =   8
               Text            =   "127"
               Top             =   7230
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.TextBox txtRefreshCycle 
               Alignment       =   1  '¾a¥k¹ï»ô
               Height          =   270
               Left            =   225
               TabIndex        =   7
               Text            =   "5000"
               Top             =   435
               Width           =   615
            End
            Begin VB.TextBox txtCycleInterval 
               Alignment       =   1  '¾a¥k¹ï»ô
               Height          =   270
               Left            =   225
               TabIndex        =   6
               Text            =   "7000"
               Top             =   3900
               Width           =   615
            End
            Begin VB.Label Label11 
               Caption         =   "Average RTT, Packet Loss% ²Î­p (default = 5 ­Óping´`Àô)"
               Height          =   375
               Left            =   240
               TabIndex        =   30
               Top             =   5700
               Visible         =   0   'False
               Width           =   4965
            End
            Begin VB.Label Label10 
               Caption         =   "continued fail as Down (default = 5 ping cycles)"
               Height          =   375
               Left            =   225
               TabIndex        =   27
               Top             =   4275
               Width           =   4965
            End
            Begin VB.Label Label22 
               Caption         =   "loss packet as fail / node / time (default = 5)"
               Height          =   375
               Left            =   225
               TabIndex        =   25
               Top             =   1920
               Width           =   4830
            End
            Begin VB.Label Label4 
               Caption         =   "response time out as fail / node (default = 1000 ms)"
               Height          =   375
               Left            =   225
               TabIndex        =   24
               Top             =   2505
               Width           =   4830
            End
            Begin VB.Label Label8 
               Caption         =   "packet count / node / time (default = 5)"
               Height          =   375
               Left            =   225
               TabIndex        =   23
               Top             =   1365
               Width           =   4830
            End
            Begin VB.Label Label2 
               Caption         =   "ping session ¼Æ / §å¦¸ / agent (default = 220 ­Ó)"
               Height          =   375
               Left            =   360
               TabIndex        =   22
               Top             =   6285
               Visible         =   0   'False
               Width           =   4830
            End
            Begin VB.Label Label5 
               Caption         =   "response time out as fail / cycle (default = 6000 ms)"
               Height          =   375
               Left            =   225
               TabIndex        =   21
               Top             =   3090
               Width           =   4830
            End
            Begin VB.Label Label6 
               Caption         =   "WFSO (default =5 ms)"
               Height          =   375
               Left            =   165
               TabIndex        =   20
               Top             =   6705
               Visible         =   0   'False
               Width           =   4830
            End
            Begin VB.Label Label1 
               Caption         =   "ping agent(default = 2 agents)"
               Height          =   375
               Left            =   225
               TabIndex        =   19
               Top             =   793
               Width           =   4830
            End
            Begin VB.Label Label3 
               Caption         =   "Delay Start Ping (default = 127 ms)"
               Height          =   375
               Left            =   915
               TabIndex        =   18
               Top             =   7005
               Visible         =   0   'False
               Width           =   4830
            End
            Begin VB.Label Label7 
               Caption         =   "Status refresh Cycle (default = 5000 ms)"
               Height          =   375
               Left            =   225
               TabIndex        =   17
               Top             =   210
               Width           =   4830
            End
            Begin VB.Label Label9 
               Caption         =   "minimum cycle interval (default = 7000 ms)"
               Height          =   375
               Left            =   225
               TabIndex        =   16
               Top             =   3675
               Width           =   4830
            End
         End
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   4725
      TabIndex        =   2
      Top             =   8415
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3510
      TabIndex        =   1
      Top             =   8415
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2295
      TabIndex        =   0
      Top             =   8415
      Width           =   1095
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
    SaveGlobalVars
    frmMain.InitAgentLedToGrayColor
    frmMain.ResizeForm
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    SaveGlobalVars
    frmMain.InitAgentLedToGrayColor
    frmMain.ResizeForm
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    '³B²z«ö¤U ctrl+tab «á¥i²¾¦Ü¤U¤@­Ó­¶ÅÒªº°Ê§@
    If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
        i = sstabOptions.Tab
        If i = (sstabOptions.Tabs - 1) Then
            '³Ì«á¤@­Ó­¶ÅÒ¡A¦]¦¹­n¦^²¾¦Ü²Ä¤@­Ó­¶ÅÒ
            sstabOptions.Tab = 0
        Else
            '»¼¼W­¶ÅÒ¯Á¤Þ­È(ªí¥Ü²¾¦Ü¤U¤@­Ó­¶ÅÒ)
            sstabOptions.Tab = sstabOptions.Tab + 1
        End If
    End If
End Sub

Private Sub Form_Load()
    
    'ªì©l¤Æ¸ê®Æ
    'txtCheckInterval = glCheckInterval
    txtContinuedFailAsDown = glContinuedFailAsDown
    txtAgentCount = glAgentCount
    txtPingCount = glPingCount
    txtThreshold = glThreshold
    txtPingTimeOutHost = glPingTimeOutHost
    'txtPingInterval = glPingInterval
    txtPingMaxBurst = glPingMaxBurst
    txtPingTimeOutBatch = glPingTimeOutBatch
    txtWaitForSingleObject = glWaitForSingleObject
    txtDelayStart = glDelayStart
    txtRefreshCycle = glRefreshCycle
    txtCycleInterval = glCycleInterval
    txtStatisticsCycle = glStatisticsCycle
    chkDebugMode.Value = IIf(glDebugMode, 1, 0)
    '±Nªí³æ¸m©ó¿Ã¹õ¤¤¥¡
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    
End Sub

Private Sub SaveGlobalVars()
    If txtAgentCount > MaxAgent Then
        MsgBox "ping agent­Ó¼Æ³Ì¤j­È" & MaxAgent, vbExclamation, MsgTitle
        txtAgentCount.SetFocus
        Exit Sub
    End If
'    If txtPingTimeOutHost > 6500 Then
'        MsgBox "³Ì¤j­È6500", vbExclamation, MsgTitle
'        txtPingTimeOutHost.SetFocus
'        Exit Sub
'    End If
'    If txtPingInterval > 6500 Then
'        MsgBox "³Ì¤j­È6500", vbExclamation, MsgTitle
'        txtPingInterval.SetFocus
'        Exit Sub
'    End If
'    If txtPingTimeOutBatch > 6500 Then
'        MsgBox "³Ì¤j­È6500", vbExclamation, MsgTitle
'        txtPingTimeOutBatch.SetFocus
'        Exit Sub
'    End If
'    If txtWaitForSingleObject > 6500 Then
'        MsgBox "³Ì¤j­È6500", vbExclamation, MsgTitle
'        txtWaitForSingleObject.SetFocus
'        Exit Sub
'    End If
    'glCheckInterval = txtCheckInterval
    glContinuedFailAsDown = txtContinuedFailAsDown
    glAgentCount = txtAgentCount
    glPingCount = txtPingCount
    glThreshold = txtThreshold
    glPingTimeOutHost = txtPingTimeOutHost
    'glPingInterval = txtPingInterval
    glPingMaxBurst = txtPingMaxBurst
    glPingTimeOutBatch = txtPingTimeOutBatch
    glWaitForSingleObject = txtWaitForSingleObject
    glDelayStart = txtDelayStart
    glRefreshCycle = txtRefreshCycle
    glCycleInterval = txtCycleInterval
    glStatisticsCycle = txtStatisticsCycle
    glDebugMode = chkDebugMode.Value
    Call SaveIniInfo
End Sub
