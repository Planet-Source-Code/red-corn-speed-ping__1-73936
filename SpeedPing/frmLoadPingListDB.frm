VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmLoadPingListDB 
   BorderStyle     =   4  '³æ½u©T©w¤u¨ãµøµ¡
   Caption         =   "Åª¨ú¸ê®Æ¤¤..."
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '¿Ã¹õ¤¤¥¡
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3780
      Top             =   1125
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   3735
      Top             =   135
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "¨ú®ø"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "§ó·s¸`ÂI¸ê®Æ, ½Ðµy«á..."
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "frmLoadPingListDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Function LoadPingListDB() As Boolean
    Dim cmd1 As ADODB.Command
    Dim rsPingList As ADODB.Recordset
    Dim i As Integer

    On Error GoTo ErrHandler
    ProgressBar1.Min = 0
    ProgressBar1.Max = MaxNodeIndex
    LoadPingListDB = False
    Set cmd1 = New Command
    cmd1.CommandType = adCmdText
    cmd1.ActiveConnection = ConnStr
    cmd1.CommandText = "delete from PingList;"
    cmd1.Execute
    'Set cmd1 = Nothing
    
    Set rsPingList = New ADODB.Recordset
    With rsPingList
        .CursorLocation = adUseClient
        .Open "PingList", ConnStr, adOpenDynamic, adLockOptimistic
        For i = 0 To MaxNodeIndex
            .AddNew
            !SN = arySN(i)
            !NodeName = aryNodeName(i)
            !Description = aryDescription(i)
            !IP = aryIPAddress(i)
            .Update
            ProgressBar1.Value = i
        Next
            
        .Close
        Set rsPingList = Nothing
    End With
    PingListDBLoaded = True
    LoadPingListDB = True
    Timer2.Enabled = True
    Exit Function
ErrHandler:
    MsgBox "¸ü¤JPingListDB¸ê®Æ®É²£¥Í¤U¦C¿ù»~:" & vbCrLf & Err.Description, vbExclamation, MsgTitle
    Timer2.Enabled = True
End Function

Private Sub Form_Load()
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    LoadPingListDB
End Sub

Private Sub Timer2_Timer()
    Timer2.Enabled = False
    Unload Me
End Sub
