VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmLoadPingList 
   BorderStyle     =   4  '³æ½u©T©w¤u¨ãµøµ¡
   Caption         =   "¼g¤J¸ê®Æ®w¤¤..."
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
      Caption         =   "¼g¤J¸ê®Æ®w¤¤, ½Ðµy«á..."
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
End
Attribute VB_Name = "frmLoadPingList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function SavePingListIntoDB() As Boolean
    Dim cnn1 As ADODB.Connection
    Dim cmd1 As ADODB.Command
    Dim rsPingList As ADODB.Recordset
    Dim i As Long
    
    '±Narray¸ü¤J¨ìDB
    On Error GoTo ErrHandler
    SavePingListIntoDB = False
    ProgressBar1.Min = 0
    ProgressBar1.Max = MaxNodeIndex
    
    '¥ý§R°£ÂÂ¸ê®Æ
    Set cnn1 = New Connection '­n¥Îconnection,§_«hcommand¦b²Ä¤G¦¸°õ¦æ®É·|¥X²{¥H¤Uªº²M°£¸ê®Æ¥¢±Ñªº©_©Ç²{¶H
    cnn1.Open ConnStr
    Set cmd1 = New Command
    cmd1.CommandType = adCmdText
    Set cmd1.ActiveConnection = cnn1
    cmd1.CommandText = "delete from PingList;"
    cmd1.Execute
    Set cmd1 = Nothing
    cnn1.Close
    Set cnn1 = Nothing
    
    'Set cmd1 = Nothing
    
    '±N¸ê®ÆÅª¤J¸ê®Æ®w¤¤
    Set rsPingList = New ADODB.Recordset
    With rsPingList
        .CursorLocation = adUseClient
        .Open "select * from PingList;", ConnStr, adOpenDynamic, adLockOptimistic
        If .RecordCount > 0 Then
            .Close
            Set rsPingList = Nothing
            MsgBox "¸ê®Æ®wpinglist²M°£ÂÂ¸ê®Æ¥¢±Ñ!", vbExclamation, MsgTitle
            Exit Function
        End If
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
    SavePingListIntoDB = True
    LoadListIntoDBOK = True
    Timer2.Enabled = True
    Exit Function
    
ErrHandler:
    MsgBox "¸ü¤JPingList¸ê®Æ¨ìDB®É²£¥Í¤U¦C¿ù»~:" & vbCrLf & Err.Description, vbExclamation, MsgTitle
    Timer2.Enabled = True
End Function

Private Sub cmdCancel_Click()
    LoadListIntoDBOK = False
    Timer2.Enabled = True
End Sub

Private Sub Form_Load()
    LoadListIntoDBOK = False
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    SavePingListIntoDB
End Sub

Private Sub Timer2_Timer()
    Timer2.Enabled = False
    Unload Me
End Sub
