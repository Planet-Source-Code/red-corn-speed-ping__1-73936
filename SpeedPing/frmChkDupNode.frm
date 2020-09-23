VERSION 5.00
Begin VB.Form frmChkDupNode 
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Âù½u©T©w¹ï¸Ü¤è¶ô
   Caption         =   "Check duplicate data..."
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5880
   ControlBox      =   0   'False
   Icon            =   "frmChkDupNode.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   405
      Top             =   2295
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3810
      Top             =   2295
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   435
      Left            =   1500
      TabIndex        =   4
      Top             =   2325
      Width           =   1755
   End
   Begin VB.PictureBox picWarn 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3660
      Picture         =   "frmChkDupNode.frx":000C
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   90
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picOK 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3030
      Picture         =   "frmChkDupNode.frx":03C2
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   90
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picArrow 
      Enabled         =   0   'False
      Height          =   375
      Left            =   2355
      Picture         =   "frmChkDupNode.frx":0778
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   90
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lblJob4 
      BackColor       =   &H80000005&
      Caption         =   "- Delete duplicate nodes"
      Enabled         =   0   'False
      Height          =   375
      Left            =   915
      TabIndex        =   11
      Top             =   1800
      Width           =   4680
   End
   Begin VB.Image imgJob4 
      Height          =   255
      Left            =   600
      Top             =   1770
      Width           =   255
   End
   Begin VB.Label lblMsg3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '³z©ú
      Caption         =   "(¦@X­Ó­«½Æ¸`ÂI)"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   3780
      TabIndex        =   10
      Top             =   1425
      Visible         =   0   'False
      Width           =   3525
   End
   Begin VB.Label lblMsg2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '³z©ú
      Caption         =   "(¦@X­Ó­«½Æ¸`ÂI)"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   3780
      TabIndex        =   9
      Top             =   1050
      Visible         =   0   'False
      Width           =   3525
   End
   Begin VB.Label lblMsg1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '³z©ú
      Caption         =   "(¦@X­Ó­«½Æ¸`ÂI)"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   3780
      TabIndex        =   8
      Top             =   675
      Visible         =   0   'False
      Width           =   3525
   End
   Begin VB.Label lblCurDoing 
      AutoSize        =   -1  'True
      BackStyle       =   0  '³z©ú
      Caption         =   "Check Items:"
      Enabled         =   0   'False
      Height          =   180
      Left            =   600
      TabIndex        =   7
      Top             =   300
      Width           =   915
   End
   Begin VB.Image imgJob2 
      Height          =   255
      Left            =   600
      Top             =   1020
      Width           =   255
   End
   Begin VB.Image imgJob3 
      Height          =   255
      Left            =   600
      Top             =   1395
      Width           =   255
   End
   Begin VB.Label lblJob3 
      BackColor       =   &H80000005&
      Caption         =   "- Duplicate IP Address"
      Enabled         =   0   'False
      Height          =   375
      Left            =   915
      TabIndex        =   6
      Top             =   1425
      Width           =   4680
   End
   Begin VB.Label lblJob2 
      BackColor       =   &H80000005&
      Caption         =   "- Duplicate Node Name"
      Enabled         =   0   'False
      Height          =   375
      Left            =   915
      TabIndex        =   5
      Top             =   1050
      Width           =   4680
   End
   Begin VB.Image imgJob1 
      Height          =   255
      Left            =   600
      Top             =   645
      Width           =   255
   End
   Begin VB.Label lblJob1 
      BackColor       =   &H80000005&
      Caption         =   "- Duplicate Node Name and IP address"
      Enabled         =   0   'False
      Height          =   375
      Left            =   915
      TabIndex        =   0
      Top             =   675
      Width           =   4680
   End
End
Attribute VB_Name = "frmChkDupNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Private Function CheckDupNode() As Boolean
    Dim cnn1 As ADODB.Connection
    Dim cmd1 As ADODB.Command
    Dim dupcount1 As Long, dupcount2 As Long, dupcount3 As Long
    Dim sqlstr As String
    
    On Error GoTo ErrHandler
    
    dupcount1 = 0: dupcount2 = 0: dupcount3 = 0
    Set cnn1 = New Connection '­n¥Îconnection,§_«hcommand¦b²Ä¤G¦¸°õ¦æ®É·|¥X²{¥H¤Uªº²M°£¸ê®Æ¥¢±Ñªº©_©Ç²{¶H
    Set cmd1 = New Command
    cnn1.Open ConnStr
    
    '§R°£ÂÂ¸ê®Æ
    cmd1.CommandType = adCmdText
    Set cmd1.ActiveConnection = cnn1
    cmd1.CommandText = "delete from DupNodeNameAndIP;"
    cmd1.Execute
    cmd1.CommandText = "delete from DupOnlyIP;"
    cmd1.Execute
    cmd1.CommandText = "delete from DupOnlyNodeName;"
    cmd1.Execute
    
    'insert­«½Æ¸ê®Æ
    cmd1.CommandText = "INSERT INTO DupNodeNameAndIP SELECT tbl1.* from PingList As tbl1, (SELECT NodeName, IP " & _
        "FROM PingList Group By NodeName, IP Having count(NodeName & IP)  > 1) As tbl2 " & _
        "where tbl1.NodeName = tbl2.NodeName And tbl1.IP = tbl2.IP;"
    cmd1.Execute
    cmd1.CommandText = "INSERT INTO DupOnlyNodeName SELECT distinct tbl1.* from PingList As tbl1, (SELECT Distinct NodeName, IP " & _
        "FROM PingList) As tbl2 " & _
        "where tbl1.NodeName =tbl2.NodeName And tbl1.IP <> tbl2.IP;"
    cmd1.Execute
    cmd1.CommandText = "INSERT INTO DupOnlyIP SELECT distinct tbl1.* from PingList As tbl1, (SELECT Distinct NodeName, IP " & _
        "FROM PingList) As tbl2 " & _
        "where tbl1.NodeName <> tbl2.NodeName And tbl1.IP = tbl2.IP order by tbl1.IP;"
    cmd1.Execute
    
    cnn1.Close
    
    Dim fs As FileSystemObject
    Dim F As TextStream
    Dim filename As String
    
    Dim rsDupNodeNameAndIP As ADODB.Recordset
    
    lblCurDoing.Enabled = True
    lblJob1.Enabled = True
    lblJob1.FontBold = True
    imgJob1.Picture = picArrow.Picture
    DoEvents

    Set rsDupNodeNameAndIP = New ADODB.Recordset
    With rsDupNodeNameAndIP
        .CursorLocation = adUseClient
        sqlstr = "SELECT * from DupNodeNameAndIP order by NodeName, SN;"
        .Open sqlstr, ConnStr, adOpenDynamic, adLockOptimistic
        If .RecordCount > 0 Then
            dupcount1 = .RecordCount
            Set fs = New FileSystemObject
            filename = App.Path & "\Duplicate - NodeName & IP.txt"
            Set F = fs.CreateTextFile(filename, True)
            .MoveFirst
            While Not .EOF
                F.WriteLine !SN & "," & !NodeName & "," & !IP
                .MoveNext
            Wend
            F.Close
            Set fs = Nothing
            
        End If
        .Close
        Set rsDupNodeNameAndIP = Nothing
    End With
    
    Dim rsDupOnlyNodeName As ADODB.Recordset
    DoEvents
    Sleep 300
    
    lblJob1.FontBold = False
    If dupcount1 > 0 Then
        imgJob1.Picture = picWarn.Picture
        lblMsg1.Caption = "(¦³" & dupcount1 & "µ§¸ê®Æ­«½Æ)"
        lblMsg1.Visible = True
    Else
        imgJob1.Picture = picOK.Picture
    End If
    
    lblJob2.Enabled = True
    lblJob2.FontBold = True
    imgJob2.Picture = picArrow.Picture
    DoEvents
    
    Set rsDupOnlyNodeName = New ADODB.Recordset
    With rsDupOnlyNodeName
        .CursorLocation = adUseClient
        sqlstr = "SELECT * from DupOnlyNodeName order by NodeName, SN;"
        .Open sqlstr, ConnStr, adOpenDynamic, adLockOptimistic
        If .RecordCount > 0 Then
            dupcount2 = .RecordCount
            Set fs = New FileSystemObject
        
            filename = App.Path & "\Duplicate - NodeName.txt"
            Set F = fs.CreateTextFile(filename, True)
            .MoveFirst
            While Not .EOF
                F.WriteLine !SN & "," & !NodeName & "," & !IP
                .MoveNext
            Wend
            F.Close
            Set fs = Nothing
        End If
        .Close
        Set rsDupOnlyNodeName = Nothing
    End With
    
    Dim rsDupOnlyIP As ADODB.Recordset
    DoEvents
    Sleep 300
    lblJob2.FontBold = False
    If dupcount2 > 0 Then
        imgJob2.Picture = picWarn.Picture
        lblMsg2.Caption = "(¦³" & dupcount2 & "µ§¸ê®Æ­«½Æ)"
        lblMsg2.Visible = True
    Else
        imgJob2.Picture = picOK.Picture
    End If
    
    lblJob3.Enabled = True
    lblJob3.FontBold = True
    imgJob3.Picture = picArrow.Picture
    DoEvents
    
    Set rsDupOnlyIP = New ADODB.Recordset
    With rsDupOnlyIP
        .CursorLocation = adUseClient
        sqlstr = "SELECT * from DupOnlyIP order by IP, SN;"
        .Open sqlstr, ConnStr, adOpenDynamic, adLockOptimistic
        If .RecordCount > 0 Then
            dupcount3 = .RecordCount
            Set fs = New FileSystemObject
        
            filename = App.Path & "\Duplicate - IP.txt"
            Set F = fs.CreateTextFile(filename, True)
            .MoveFirst
            While Not .EOF
                F.WriteLine !SN & "," & !NodeName & "," & !IP
                .MoveNext
            Wend
            F.Close
            Set fs = Nothing
        End If
        .Close
        Set rsDupOnlyIP = Nothing
    End With
    
    DoEvents
    Sleep 300
    lblJob3.FontBold = False
    If dupcount3 > 0 Then
        imgJob3.Picture = picWarn.Picture
        lblMsg3.Caption = "(¦³" & dupcount3 & "µ§¸ê®Æ­«½Æ)"
        lblMsg3.Visible = True
    Else
        imgJob3.Picture = picOK.Picture
    End If
    
    lblJob4.Enabled = True
    lblJob4.FontBold = True
    imgJob4.Picture = picArrow.Picture
    DoEvents
    
    
    If dupcount1 > 0 Or dupcount2 > 0 Or dupcount3 > 0 Then
        '§R°£­«½Æªº¸ê®Æ
        cnn1.Open ConnStr
        cmd1.CommandType = adCmdText
        Set cmd1.ActiveConnection = cnn1
        cmd1.CommandText = "delete tbl1.* from PingList As tbl1, DupNodeNameAndIP as tbl2 " & _
            "where tbl1.SN = tbl2.SN;"
        cmd1.Execute
        cmd1.CommandText = "delete tbl1.* from PingList As tbl1, DupOnlyNodeName as tbl2 " & _
            "where tbl1.SN = tbl2.SN;"
        cmd1.Execute
        cmd1.CommandText = "delete tbl1.* from PingList As tbl1, DupOnlyIP as tbl2 " & _
            "where tbl1.SN = tbl2.SN;"
        cmd1.Execute
        cnn1.Close
        
        DoEvents
        Sleep 300
        lblJob4.FontBold = False
        imgJob4.Picture = picOK.Picture
    
        CheckDupNodeOK = True
    Else
        CheckDupNodeOK = True
        Timer2.Enabled = True
    End If
    
    Me.Caption = "ÀË¬d§¹¦¨!"
    cmdOK.Enabled = True
    
    Exit Function
ErrHandler:
    MsgBox "°õ¦æ¸ê®Æ­«½ÆªºÀË¬d®Éµo¥Í¿ù»~:" & vbCrLf & Err.Description, vbExclamation, MsgTitle
    cmdOK.Enabled = True
End Function

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    CheckDupNodeOK = False
    Call CheckDupNode
End Sub

Private Sub Timer2_Timer()
    Unload Me
End Sub
