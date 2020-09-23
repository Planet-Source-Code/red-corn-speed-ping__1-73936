VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLog 
   Caption         =   "log"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13515
   Icon            =   "frmLog.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   13515
   StartUpPosition =   2  '¿Ã¹õ¤¤¥¡
   Begin VB.PictureBox Picture1 
      Align           =   1  '¹ï»ôªí³æ¤W¤è
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   1005
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   13515
      TabIndex        =   1
      Top             =   0
      Width           =   13515
      Begin VB.PictureBox Picture2 
         Appearance      =   0  '¥­­±
         BorderStyle     =   0  '¨S¦³®Ø½u
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   105
         ScaleHeight     =   465
         ScaleWidth      =   5160
         TabIndex        =   14
         Top             =   420
         Width           =   5160
         Begin VB.CommandButton cmdClearLog 
            Caption         =   "Clear Event Log"
            Height          =   315
            Left            =   2835
            TabIndex        =   16
            Top             =   0
            Width           =   1695
         End
         Begin VB.TextBox txtLogKeepDays 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Height          =   285
            Left            =   1305
            TabIndex        =   15
            Text            =   "7"
            Top             =   10
            Width           =   690
         End
         Begin VB.Label Label6 
            Caption         =   "Reserve event in:"
            Height          =   285
            Left            =   0
            TabIndex        =   18
            Top             =   55
            Width           =   1500
         End
         Begin VB.Label Label7 
            Caption         =   "days -->"
            Height          =   285
            Left            =   2055
            TabIndex        =   17
            Top             =   60
            Width           =   750
         End
      End
      Begin VB.TextBox txtDesc3 
         Appearance      =   0  '¥­­±
         BackColor       =   &H0080FFFF&
         Height          =   285
         Left            =   11835
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   90
         Width           =   1545
      End
      Begin VB.TextBox txtDesc2 
         Appearance      =   0  '¥­­±
         BackColor       =   &H0080FFFF&
         Height          =   285
         Left            =   9105
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   90
         Width           =   1545
      End
      Begin VB.TextBox txtDesc1 
         Appearance      =   0  '¥­­±
         BackColor       =   &H0080FFFF&
         Height          =   285
         Left            =   6405
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   90
         Width           =   1545
      End
      Begin VB.TextBox txtIP 
         Appearance      =   0  '¥­­±
         BackColor       =   &H0080FFFF&
         Height          =   285
         Left            =   3690
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   90
         Width           =   1545
      End
      Begin VB.TextBox txtNodeName 
         Appearance      =   0  '¥­­±
         BackColor       =   &H0080FFFF&
         Height          =   285
         Left            =   1050
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   90
         Width           =   1545
      End
      Begin VB.Label Label5 
         Alignment       =   1  '¾a¥k¹ï»ô
         Caption         =   "Description 3:"
         Height          =   285
         Left            =   10785
         TabIndex        =   13
         Top             =   135
         Width           =   1020
      End
      Begin VB.Label Label4 
         Alignment       =   1  '¾a¥k¹ï»ô
         Caption         =   "Description 2:"
         Height          =   285
         Left            =   8055
         TabIndex        =   11
         Top             =   135
         Width           =   1020
      End
      Begin VB.Label Label3 
         Alignment       =   1  '¾a¥k¹ï»ô
         Caption         =   "Description 1:"
         Height          =   285
         Left            =   5355
         TabIndex        =   8
         Top             =   135
         Width           =   1020
      End
      Begin VB.Label Label2 
         Alignment       =   1  '¾a¥k¹ï»ô
         Caption         =   "IP Address:"
         Height          =   285
         Left            =   2700
         TabIndex        =   6
         Top             =   135
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "Node Name:"
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   135
         Width           =   930
      End
   End
   Begin MSComctlLib.ListView lvLog1 
      Height          =   2610
      Left            =   90
      TabIndex        =   0
      Top             =   1110
      Width           =   6870
      _ExtentX        =   12118
      _ExtentY        =   4604
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      PictureAlignment=   4
      _Version        =   393217
      SmallIcons      =   "imlSmall"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "·s²Ó©úÅé"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Event"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSortIcon 
      Left            =   8235
      Top             =   1125
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   8
      ImageHeight     =   7
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLog.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLog.frx":03DC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar statusbar 
      Align           =   2  '¹ï»ôªí³æ¤U¤è
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   6240
      Width           =   13515
      _ExtentX        =   23839
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   23310
            Text            =   "°T®§"
            TextSave        =   "°T®§"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSmall 
      Left            =   7650
      Top             =   1935
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLog.frx":04AE
            Key             =   "green"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLog.frx":0EC0
            Key             =   "red"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLog.frx":18D2
            Key             =   "yellow"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLog.frx":22E4
            Key             =   "gray"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvLog2 
      Height          =   2610
      Left            =   2700
      TabIndex        =   9
      Top             =   3420
      Width           =   6870
      _ExtentX        =   12118
      _ExtentY        =   4604
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      PictureAlignment=   4
      _Version        =   393217
      SmallIcons      =   "imlSmall"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "·s²Ó©úÅé"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "SN"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Node Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "IP Address"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Description 1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Description 2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Description 3"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Event"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const WM_SETREDRAW = &HB
Private Const LVM_FIRST As Long = &H1000
Private Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
Private Const LVSCW_AUTOSIZE As Long = -1
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2
Private Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long
   
Private Const WM_USER = &H400
Private Const EM_SETREADONLY = (WM_USER + 31)

Private Type RECT
   Left    As Long
   Top     As Long
   Right   As Long
   Bottom  As Long
End Type
Private Declare Function GetClientRect& Lib "user32" _
                            (ByVal hWnd&, Rct As RECT)

Private ShowNodeName As Boolean

Private ListViewSortOrder() As Integer

Private Sub cmdClearLog_Click()

    Dim response, msg, Style, title
    Dim LogKeepDays As Integer
    
    On Error GoTo ErrHandler
    If Not IsNumeric(txtLogKeepDays.Text) Then
        txtLogKeepDays.SetFocus
        Exit Sub
    End If
    LogKeepDays = txtLogKeepDays.Text
    
    If LogKeepDays < 0 Then
        txtLogKeepDays.SetFocus
        Exit Sub
    End If
    title = "²M°£¨Æ¥ó°O¿ý"  ' ©w¸q¼ÐÃD
    
    Style = vbYesNo + vbCritical + vbDefaultButton1   ' ©w¸q«ö¶s
    msg = "±z½T©w­n²M°£¶W¹L " & LogKeepDays & " ¤Ñªº¨Æ¥ó°O¿ý¸ê®Æ¶Ü?"   ' ©w¸q°T®§
    response = MsgBox(msg, Style, title)
    If response = vbNo Then   ' ­Y¨Ï¥ÎªÌ«ö¤U [§_]
        Exit Sub
    End If

    Dim cnn1 As ADODB.Connection
    Dim cmd1 As ADODB.Command
    
    Set cnn1 = New Connection '­n¥Îconnection,§_«hcommand¦b²Ä¤G¦¸°õ¦æ®É·|¥X²{¥H¤Uªº²M°£¸ê®Æ¥¢±Ñªº©_©Ç²{¶H
    Set cmd1 = New Command
    cnn1.Open ConnStr
    
    '§R°£ÂÂ¸ê®Æ
    cmd1.CommandType = adCmdText
    Set cmd1.ActiveConnection = cnn1
    ' '¤µ¤Ñºâ¤@¤Ñ
    cmd1.CommandText = "delete * from UpDown where DateDiff (""d"", UpDown!LogTime,""" & Format(Date, "yyyy/mm/dd") & """ )  >= " & LogKeepDays & ";"
    cmd1.Execute
    cnn1.Close
    Call LoadLogData
    Exit Sub
ErrHandler:
    MsgBox "°õ¦æ¨Æ¥ó°O¿ý²M°£®Éµo¥Í¿ù»~:" & vbCrLf & Err.Description, vbExclamation, MsgTitle
    
End Sub

Private Sub Form_Load()
    lvLog1.Left = 45
    lvLog2.Left = 45

End Sub
Public Sub SetNodeName(NodeName As String)
    txtNodeName.Text = NodeName
    '¥Î¥H¤Uªº¤èªkµL®Ä,±Nlocked = true§Y¥i
    'SendMessage txtNodeName.hWnd, EM_SETREADONLY, True, 0&
End Sub
Public Sub SetIP(IP As String)
    txtIP.Text = IP
    'SendMessage txtIP.hWnd, EM_SETREADONLY, True, 0&
End Sub
Public Sub SetDesc(Desc1 As String, Optional Desc2 As String = "", Optional Desc3 As String = "")
    txtDesc1.Text = Desc1
    txtDesc2.Text = Desc2
    txtDesc3.Text = Desc3
    'SendMessage txtDesc.hWnd, EM_SETREADONLY, True, 0&
End Sub

Public Sub SetShowNodeName(Optional b As Boolean = True)
    ShowNodeName = b
    If ShowNodeName Then
        lvLog1.Visible = False
        lvLog2.Visible = True 'lvLog2¦³Åã¥Ü:SN, Nodename, IP addr, Description 1,2,3
    Else
        lvLog1.Visible = True
        lvLog2.Visible = False
    End If
End Sub
Private Sub Form_Resize()
    Dim clientrect As RECT
    Dim lvlog As ListView
    On Error GoTo ErrHandler
    If Me.WindowState = vbMinimized Then Exit Sub
    GetClientRect Me.hWnd, clientrect
    If ShowNodeName Then
        Set lvlog = lvLog2
        Picture1.Height = 780
        Picture2.Visible = True
    Else
        Set lvlog = lvLog1
        Picture1.Height = 420 'ÁôÂÃ²M°£logªº¥\¯à
        Picture2.Visible = False
    End If
    'ª`·N:¤£­n¥Îme.Height,¤@¥¹¤Á´«classic style / xp style¤£¦Pªºcontrol®É,¶ZÂ÷·|¦³»~®t
    lvLog1.Top = 10 + Picture1.Height
    lvLog2.Top = 10 + Picture1.Height
    
    lvlog.Height = (clientrect.Bottom - clientrect.Top) * Screen.TwipsPerPixelY - 20 - statusbar.Height - Picture1.Height
    lvlog.Width = (clientrect.Right - clientrect.Left) * Screen.TwipsPerPixelX - 90
    
    AdjustColWidth lvlog
    Exit Sub
ErrHandler:
End Sub
Public Sub LoadLogData(Optional NodeName As String = "")
    
    Dim rsUpDown As ADODB.Recordset
    Dim x As Long
    Dim sqlstr As String
    Dim i As Integer
    Dim itemx As ListItem
    Dim lvlog As ListView
    Dim sbuffer As String
    Dim maxnumlen As Integer
    
    '¶}©lÀx¦s¸ê®Æ
    On Error GoTo ErrHandler
    
    maxnumlen = Len(CStr(NumOfPingNode))
    sbuffer = Space(maxnumlen)
    
    Set rsUpDown = New ADODB.Recordset
    i = 0
    statusbar.Panels(1).Text = ""
    
    If ShowNodeName Then
        Set lvlog = lvLog2
    Else
        Set lvlog = lvLog1
    End If
    With rsUpDown
        .CursorLocation = adUseClient
        
        '¦¨¥æ¶q
        If ShowNodeName Then
'            If Not PingListDBLoaded Then
'                frmLoadPingListDB.Show vbModal, Me
'                If Not PingListDBLoaded Then
'                    Exit Sub
'                End If
'            End If
            'sqlstr = "select * from UpDown order by LogTime;"
            sqlstr = "SELECT UpDown.*, PingList.SN,PingList.IP,PingList.Route1,PingList.Route2,PingList.Route3 " & _
                "FROM PingList INNER JOIN UpDown ON PingList.NodeName = UpDown.NodeName;"

        Else
            sqlstr = "select * from UpDown where NodeName = """ & NodeName & """ order by LogTime;"
        End If
  
        .Open sqlstr, ConnStr, adOpenDynamic, adLockOptimistic

        If .RecordCount = 0 Then
            x = SendMessage(lvlog.hWnd, WM_SETREDRAW, 0, 0)
            lvlog.ListItems.Clear
            x = SendMessage(lvlog.hWnd, WM_SETREDRAW, 1, 0)
            statusbar.Panels(1).Text = "¬dµL¸ê®Æ!"
        Else
            statusbar.Panels(1).Text = "¦@" & .RecordCount & "µ§°O¿ý!"
            .MoveFirst
            x = SendMessage(lvlog.hWnd, WM_SETREDRAW, 0, 0)
            lvlog.ListItems.Clear
            lvlog.ColumnHeaders(1).Text = ""
            lvlog.ColumnHeaders(1).Width = 300
            If ShowNodeName Then
                While Not .EOF
                    i = i + 1
                    If !Event = 1 Then
                        Set itemx = lvlog.ListItems.Add(, "#" & i, " ", , "green")
                        itemx.SubItems(8) = "Up"
                    Else
                        Set itemx = lvlog.ListItems.Add(, "#" & i, "", , "red")
                        itemx.SubItems(8) = "Down"
                    End If
                    itemx.SubItems(1) = Format(!LogTime, "yyyy/mm/dd Hh:Nn:Ss")
                    itemx.SubItems(2) = Right(sbuffer & !SN, maxnumlen)
                    itemx.SubItems(3) = !NodeName
                    itemx.SubItems(4) = !IP
                    itemx.SubItems(5) = !Route1
                    itemx.SubItems(6) = !Route2
                    itemx.SubItems(7) = !Route3
                    .MoveNext
                Wend
            Else
                While Not .EOF
                    i = i + 1
                    If !Event = 1 Then
                        Set itemx = lvlog.ListItems.Add(, "#" & i, " ", , "green")
                        itemx.SubItems(2) = "Up"
                    Else
                        Set itemx = lvlog.ListItems.Add(, "#" & i, "", , "red")
                        itemx.SubItems(2) = "Down"
                    End If
                    itemx.SubItems(1) = Format(!LogTime, "yyyy/mm/dd Hh:Nn:Ss")
                    .MoveNext
                Wend
            
            End If
            x = SendMessage(lvlog.hWnd, WM_SETREDRAW, 1, 0)
            AdjustColWidth lvlog
            ReDim ListViewSortOrder(1 To lvlog.ColumnHeaders.Count)
            Call InitListViewSort(lvlog)
        End If
        .Close
        Set rsUpDown = Nothing
    End With
  
    Exit Sub
    
ErrHandler:
    MsgBox "Åª¨ú¸ê®Æ®É²£¥Í¤U¦C¿ù»~:" & vbCrLf & Err.Description, vbExclamation, MsgTitle
End Sub
Private Sub InitListViewSort(lv As ListView)
    Dim col As ColumnHeader
    Set col = lv.ColumnHeaders(2)     'sort on first column
    lv.ColumnHeaderIcons = imlSortIcon
    ListViewSortOrder(2) = lvwDescending      'will get flipped to ascending
    ListViewColumnClick lv, col                'click the column heading
    lv.ListItems(1).EnsureVisible     'make sure first one is visible
    lv.ListItems(1).Selected = True   'and selected
End Sub



Private Sub lvLog1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListViewColumnClick lvLog1, ColumnHeader
End Sub
Private Sub ListViewColumnClick(lv As ListView, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
' °õ¦æ±Æ§Ç
' 1 is the first column
    Dim i As Integer
    With lv
        If .ListItems.Count = 0 Then Exit Sub
        For i = 1 To .ColumnHeaders.Count
            If i = ColumnHeader.Index Then
                ListViewSortOrder(i) = FlipSort(ListViewSortOrder(ColumnHeader.Index))
            Else
                ListViewSortOrder(i) = lvwDescending '¨ä¥¦ªº³]¬°Descending,¤U¤@¦¸click®ÉÅÜ¬°Ascending
            End If
        Next
        .SortOrder = ListViewSortOrder(ColumnHeader.Index)
        .SortKey = ColumnHeader.Index - 1
        .Sorted = True
        DoEvents
        'If chkAutoSort.Value <> 1 Then
            .Sorted = False
        'End If
        'Show column icon
        ShowListViewSortIcon lv
        
'        If Not .SelectedItem Is Nothing Then
'            .SelectedItem.EnsureVisible
'        End If
        .ListItems(1).EnsureVisible     'make sure first one is visible
        .ListItems(1).Selected = True   'and selected
        
    End With
End Sub



Private Sub lvLog2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    ListViewColumnClick lvLog2, ColumnHeader
End Sub


Private Sub lvLog2_DblClick()
'lvLog2¬O©Ò¦³nodeªºlog,©Ò¥H¦³double click,¥i¥H¦A¥s¥X­Ó§Onodeªº©ú²Ó
    Dim itemx As ListItem
    Dim NodeName As String
    Dim frm As frmLog
    
    On Error GoTo ErrHandler
    
    If lvLog2.SelectedItem Is Nothing Then
        
         Exit Sub
    End If
    Set itemx = lvLog2.SelectedItem
    
    NodeName = itemx.SubItems(3)
    If NodeName = "" Then Exit Sub

    Set frm = New frmLog '¦¹®É¤w°õ¦æform Load event
    frm.Caption = "log - " & NodeName
    frm.SetNodeName NodeName
    frm.SetIP itemx.SubItems(4)
    frm.SetDesc itemx.SubItems(5), itemx.SubItems(6), itemx.SubItems(7)
    frm.SetShowNodeName False
    frm.LoadLogData NodeName
    frm.Show vbModeless
    Exit Sub
ErrHandler:
    MsgBox "Error!" & vbCrLf & Err.Description, vbExclamation, MsgTitle
End Sub

Private Sub lvLog2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Set lvLog2.SelectedItem = lvLog2.HitTest(x, y)
End Sub
Public Sub AdjustColWidth(lv As ListView, Optional HideCol As Integer = -1)
    'Size each column based on the maximum of
  'EITHER the column header text width, or,
  'if the items below it are wider, the
  'widest list item in the column.
  '
  'The last column is always resized to occupy
  'the remaining width in the control.
    Dim startcol As Long
    Dim col2adjust As Long
    If lv.View = lvwReport Then
        
    '¦Û°Ê½Õ¾ãÄæ¼e
        'lv.ColumnHeaders(1).Width = 300
        If lv.ColumnHeaders(1).Text = "" Then
            startcol = 1
        Else
            startcol = 0
        End If
        For col2adjust = startcol To lv.ColumnHeaders.Count - 1
        
            If col2adjust = HideCol Then
                lv.ColumnHeaders(col2adjust).Width = 0
            Else
                Call SendMessage(lv.hWnd, _
                     LVM_SETCOLUMNWIDTH, _
                     col2adjust, _
                     ByVal LVSCW_AUTOSIZE_USEHEADER)
            End If
        Next
        
    End If
End Sub

