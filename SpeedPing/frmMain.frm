VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Ping"
   ClientHeight    =   5625
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15240
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   15240
   WindowState     =   2  '³Ì¤j¤Æ
   Begin VB.Timer tmrCheckCollectorAlive 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9525
      Top             =   4545
   End
   Begin VB.Timer tmrChkAgentReady 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   10245
      Top             =   3960
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Test2"
      Height          =   330
      Left            =   10095
      TabIndex        =   26
      Top             =   3405
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Test3"
      Height          =   330
      Left            =   11250
      TabIndex        =   25
      Top             =   3405
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Timer tmrLog 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   9555
      Top             =   3960
   End
   Begin VB.PictureBox picToolbar 
      Align           =   1  '¹ï»ôªí³æ¤W¤è
      Appearance      =   0  '¥­­±
      BorderStyle     =   0  '¨S¦³®Ø½u
      ForeColor       =   &H80000008&
      Height          =   1845
      Left            =   0
      ScaleHeight     =   1845
      ScaleWidth      =   15240
      TabIndex        =   21
      Top             =   0
      Width           =   15240
      Begin VB.CommandButton cmdAbout 
         Caption         =   "About"
         Height          =   330
         Left            =   7245
         TabIndex        =   5
         Top             =   60
         Width           =   1080
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  '¥­­±
         BorderStyle     =   0  '¨S¦³®Ø½u
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   135
         ScaleHeight     =   330
         ScaleWidth      =   12840
         TabIndex        =   270
         Top             =   1455
         Width           =   12840
         Begin VB.CommandButton cmdOpenEventLog 
            Caption         =   "Open Event Log"
            Height          =   330
            Left            =   10200
            TabIndex        =   12
            Top             =   0
            Width           =   1485
         End
         Begin VB.CommandButton cmdExcel 
            Caption         =   "Excel"
            Height          =   330
            Left            =   9120
            TabIndex        =   11
            Top             =   0
            Width           =   1080
         End
         Begin VB.TextBox txtFind 
            Height          =   285
            Left            =   2220
            TabIndex        =   6
            Top             =   30
            Width           =   1545
         End
         Begin VB.CommandButton cmdFindFirst 
            Caption         =   "Find First"
            Height          =   330
            Left            =   3840
            TabIndex        =   7
            Top             =   0
            Width           =   1080
         End
         Begin VB.CommandButton cmdFindNext 
            Caption         =   "Find Next"
            Height          =   330
            Left            =   4920
            TabIndex        =   8
            Top             =   0
            Width           =   1080
         End
         Begin VB.CommandButton cmdGoTo 
            Caption         =   "Go To"
            Height          =   330
            Left            =   7575
            TabIndex        =   10
            Top             =   0
            Width           =   1080
         End
         Begin VB.TextBox txtGoTo 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Height          =   285
            Left            =   6930
            TabIndex        =   9
            Top             =   30
            Width           =   585
         End
         Begin VB.PictureBox picPingCollector 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   705
            Picture         =   "frmMain.frx":08CA
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   271
            TabStop         =   0   'False
            Top             =   75
            Width           =   240
         End
         Begin VB.Label Label1 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "Find Node:"
            Height          =   195
            Left            =   1350
            TabIndex        =   274
            Top             =   75
            Width           =   900
         End
         Begin VB.Label Label2 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "Go to #:"
            Height          =   195
            Left            =   6225
            TabIndex        =   273
            Top             =   75
            Width           =   720
         End
         Begin VB.Label Label3 
            Caption         =   "Collector:"
            Height          =   195
            Left            =   0
            TabIndex        =   272
            Top             =   75
            Width           =   720
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '¥­­±
         BorderStyle     =   0  '¨S¦³®Ø½u
         ForeColor       =   &H80000008&
         Height          =   990
         Left            =   60
         ScaleHeight     =   990
         ScaleWidth      =   17280
         TabIndex        =   29
         Top             =   450
         Width           =   17280
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   119
            Left            =   16875
            Picture         =   "frmMain.frx":12CC
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   267
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   118
            Left            =   16875
            Picture         =   "frmMain.frx":1CCE
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   266
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   117
            Left            =   16590
            Picture         =   "frmMain.frx":26D0
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   263
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   116
            Left            =   16590
            Picture         =   "frmMain.frx":30D2
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   262
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   115
            Left            =   16305
            Picture         =   "frmMain.frx":3AD4
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   259
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   114
            Left            =   16305
            Picture         =   "frmMain.frx":44D6
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   258
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   113
            Left            =   16020
            Picture         =   "frmMain.frx":4ED8
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   255
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   112
            Left            =   16020
            Picture         =   "frmMain.frx":58DA
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   254
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   111
            Left            =   15735
            Picture         =   "frmMain.frx":62DC
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   251
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   110
            Left            =   15735
            Picture         =   "frmMain.frx":6CDE
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   250
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   109
            Left            =   15450
            Picture         =   "frmMain.frx":76E0
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   247
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   108
            Left            =   15450
            Picture         =   "frmMain.frx":80E2
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   246
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   107
            Left            =   15165
            Picture         =   "frmMain.frx":8AE4
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   243
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   106
            Left            =   15165
            Picture         =   "frmMain.frx":94E6
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   242
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   105
            Left            =   14880
            Picture         =   "frmMain.frx":9EE8
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   239
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   104
            Left            =   14880
            Picture         =   "frmMain.frx":A8EA
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   238
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   103
            Left            =   14595
            Picture         =   "frmMain.frx":B2EC
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   235
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   102
            Left            =   14595
            Picture         =   "frmMain.frx":BCEE
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   234
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   101
            Left            =   14310
            Picture         =   "frmMain.frx":C6F0
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   231
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   100
            Left            =   14310
            Picture         =   "frmMain.frx":D0F2
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   230
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   99
            Left            =   14025
            Picture         =   "frmMain.frx":DAF4
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   227
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   98
            Left            =   14025
            Picture         =   "frmMain.frx":E4F6
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   226
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   97
            Left            =   13740
            Picture         =   "frmMain.frx":EEF8
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   223
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   96
            Left            =   13740
            Picture         =   "frmMain.frx":F8FA
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   222
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   95
            Left            =   13455
            Picture         =   "frmMain.frx":102FC
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   219
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   94
            Left            =   13455
            Picture         =   "frmMain.frx":10CFE
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   218
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   93
            Left            =   13170
            Picture         =   "frmMain.frx":11700
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   215
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   92
            Left            =   13170
            Picture         =   "frmMain.frx":12102
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   214
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   91
            Left            =   12885
            Picture         =   "frmMain.frx":12B04
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   211
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   90
            Left            =   12885
            Picture         =   "frmMain.frx":13506
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   210
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   89
            Left            =   12600
            Picture         =   "frmMain.frx":13F08
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   207
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   88
            Left            =   12600
            Picture         =   "frmMain.frx":1490A
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   206
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   87
            Left            =   12315
            Picture         =   "frmMain.frx":1530C
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   203
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   86
            Left            =   12315
            Picture         =   "frmMain.frx":15D0E
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   202
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   85
            Left            =   12030
            Picture         =   "frmMain.frx":16710
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   199
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   84
            Left            =   12030
            Picture         =   "frmMain.frx":17112
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   198
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   83
            Left            =   11745
            Picture         =   "frmMain.frx":17B14
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   195
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   82
            Left            =   11745
            Picture         =   "frmMain.frx":18516
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   194
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   81
            Left            =   11460
            Picture         =   "frmMain.frx":18F18
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   191
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   80
            Left            =   11460
            Picture         =   "frmMain.frx":1991A
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   190
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   79
            Left            =   11175
            Picture         =   "frmMain.frx":1A31C
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   187
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   78
            Left            =   11175
            Picture         =   "frmMain.frx":1AD1E
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   186
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   77
            Left            =   10890
            Picture         =   "frmMain.frx":1B720
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   183
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   76
            Left            =   10890
            Picture         =   "frmMain.frx":1C122
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   182
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   75
            Left            =   10605
            Picture         =   "frmMain.frx":1CB24
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   179
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   74
            Left            =   10605
            Picture         =   "frmMain.frx":1D526
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   178
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   73
            Left            =   10320
            Picture         =   "frmMain.frx":1DF28
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   175
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   72
            Left            =   10320
            Picture         =   "frmMain.frx":1E92A
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   174
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   71
            Left            =   10035
            Picture         =   "frmMain.frx":1F32C
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   171
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   70
            Left            =   10035
            Picture         =   "frmMain.frx":1FD2E
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   170
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   69
            Left            =   9750
            Picture         =   "frmMain.frx":20730
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   167
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   68
            Left            =   9750
            Picture         =   "frmMain.frx":21132
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   166
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   67
            Left            =   9465
            Picture         =   "frmMain.frx":21B34
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   163
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   66
            Left            =   9465
            Picture         =   "frmMain.frx":22536
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   162
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   65
            Left            =   9180
            Picture         =   "frmMain.frx":22F38
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   159
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   64
            Left            =   9180
            Picture         =   "frmMain.frx":2393A
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   158
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   63
            Left            =   8895
            Picture         =   "frmMain.frx":2433C
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   155
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   62
            Left            =   8895
            Picture         =   "frmMain.frx":24D3E
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   154
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   61
            Left            =   8610
            Picture         =   "frmMain.frx":25740
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   151
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   60
            Left            =   8610
            Picture         =   "frmMain.frx":26142
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   150
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   59
            Left            =   8325
            Picture         =   "frmMain.frx":26B44
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   147
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   58
            Left            =   8325
            Picture         =   "frmMain.frx":27546
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   146
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   57
            Left            =   8040
            Picture         =   "frmMain.frx":27F48
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   143
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   56
            Left            =   8040
            Picture         =   "frmMain.frx":2894A
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   142
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   55
            Left            =   7755
            Picture         =   "frmMain.frx":2934C
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   139
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   54
            Left            =   7755
            Picture         =   "frmMain.frx":29D4E
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   138
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   53
            Left            =   7470
            Picture         =   "frmMain.frx":2A750
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   135
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   52
            Left            =   7470
            Picture         =   "frmMain.frx":2B152
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   134
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   51
            Left            =   7185
            Picture         =   "frmMain.frx":2BB54
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   131
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   50
            Left            =   7185
            Picture         =   "frmMain.frx":2C556
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   130
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   49
            Left            =   6900
            Picture         =   "frmMain.frx":2CF58
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   127
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   48
            Left            =   6900
            Picture         =   "frmMain.frx":2D95A
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   126
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   47
            Left            =   6615
            Picture         =   "frmMain.frx":2E35C
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   123
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   46
            Left            =   6615
            Picture         =   "frmMain.frx":2ED5E
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   122
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   45
            Left            =   6330
            Picture         =   "frmMain.frx":2F760
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   119
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   44
            Left            =   6330
            Picture         =   "frmMain.frx":30162
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   118
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   43
            Left            =   6045
            Picture         =   "frmMain.frx":30B64
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   115
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   42
            Left            =   6045
            Picture         =   "frmMain.frx":31566
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   114
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   41
            Left            =   5760
            Picture         =   "frmMain.frx":31F68
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   111
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   40
            Left            =   5760
            Picture         =   "frmMain.frx":3296A
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   110
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   39
            Left            =   5475
            Picture         =   "frmMain.frx":3336C
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   107
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   38
            Left            =   5475
            Picture         =   "frmMain.frx":33D6E
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   106
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   37
            Left            =   5190
            Picture         =   "frmMain.frx":34770
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   103
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   36
            Left            =   5190
            Picture         =   "frmMain.frx":35172
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   102
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   35
            Left            =   4905
            Picture         =   "frmMain.frx":35B74
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   99
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   34
            Left            =   4905
            Picture         =   "frmMain.frx":36576
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   98
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   33
            Left            =   4620
            Picture         =   "frmMain.frx":36F78
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   95
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   32
            Left            =   4620
            Picture         =   "frmMain.frx":3797A
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   94
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   31
            Left            =   4335
            Picture         =   "frmMain.frx":3837C
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   91
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   30
            Left            =   4335
            Picture         =   "frmMain.frx":38D7E
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   90
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   29
            Left            =   4050
            Picture         =   "frmMain.frx":39780
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   87
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   28
            Left            =   4050
            Picture         =   "frmMain.frx":3A182
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   86
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   27
            Left            =   3765
            Picture         =   "frmMain.frx":3AB84
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   83
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   26
            Left            =   3765
            Picture         =   "frmMain.frx":3B586
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   82
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   25
            Left            =   3480
            Picture         =   "frmMain.frx":3BF88
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   79
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   24
            Left            =   3480
            Picture         =   "frmMain.frx":3C98A
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   78
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   23
            Left            =   3195
            Picture         =   "frmMain.frx":3D38C
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   75
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   22
            Left            =   3195
            Picture         =   "frmMain.frx":3DD8E
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   74
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   21
            Left            =   2910
            Picture         =   "frmMain.frx":3E790
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   20
            Left            =   2910
            Picture         =   "frmMain.frx":3F192
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   19
            Left            =   2625
            Picture         =   "frmMain.frx":3FB94
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   18
            Left            =   2625
            Picture         =   "frmMain.frx":40596
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   17
            Left            =   2340
            Picture         =   "frmMain.frx":40F98
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   16
            Left            =   2340
            Picture         =   "frmMain.frx":4199A
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   15
            Left            =   2055
            Picture         =   "frmMain.frx":4239C
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   14
            Left            =   2055
            Picture         =   "frmMain.frx":42D9E
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   13
            Left            =   1770
            Picture         =   "frmMain.frx":437A0
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   12
            Left            =   1770
            Picture         =   "frmMain.frx":441A2
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   11
            Left            =   1485
            Picture         =   "frmMain.frx":44BA4
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   10
            Left            =   1485
            Picture         =   "frmMain.frx":455A6
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   9
            Left            =   1200
            Picture         =   "frmMain.frx":45FA8
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   8
            Left            =   1200
            Picture         =   "frmMain.frx":469AA
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   7
            Left            =   915
            Picture         =   "frmMain.frx":473AC
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   6
            Left            =   915
            Picture         =   "frmMain.frx":47DAE
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   5
            Left            =   630
            Picture         =   "frmMain.frx":487B0
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   4
            Left            =   630
            Picture         =   "frmMain.frx":491B2
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   3
            Left            =   345
            Picture         =   "frmMain.frx":49BB4
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   2
            Left            =   345
            Picture         =   "frmMain.frx":4A5B6
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   0
            Left            =   60
            Picture         =   "frmMain.frx":4AFB8
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   1
            Left            =   60
            Picture         =   "frmMain.frx":4B9BA
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   705
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   119
            Left            =   16875
            TabIndex        =   269
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   118
            Left            =   16875
            TabIndex        =   268
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   117
            Left            =   16590
            TabIndex        =   265
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   116
            Left            =   16590
            TabIndex        =   264
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   115
            Left            =   16305
            TabIndex        =   261
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   114
            Left            =   16305
            TabIndex        =   260
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   113
            Left            =   16020
            TabIndex        =   257
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   112
            Left            =   16020
            TabIndex        =   256
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   111
            Left            =   15735
            TabIndex        =   253
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   110
            Left            =   15735
            TabIndex        =   252
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   109
            Left            =   15450
            TabIndex        =   249
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   108
            Left            =   15450
            TabIndex        =   248
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   107
            Left            =   15165
            TabIndex        =   245
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   106
            Left            =   15165
            TabIndex        =   244
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   105
            Left            =   14880
            TabIndex        =   241
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   104
            Left            =   14880
            TabIndex        =   240
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   103
            Left            =   14595
            TabIndex        =   237
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   102
            Left            =   14595
            TabIndex        =   236
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   101
            Left            =   14310
            TabIndex        =   233
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   100
            Left            =   14310
            TabIndex        =   232
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   99
            Left            =   14025
            TabIndex        =   229
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   98
            Left            =   14025
            TabIndex        =   228
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   97
            Left            =   13740
            TabIndex        =   225
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   96
            Left            =   13740
            TabIndex        =   224
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   95
            Left            =   13455
            TabIndex        =   221
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   94
            Left            =   13455
            TabIndex        =   220
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   93
            Left            =   13170
            TabIndex        =   217
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   92
            Left            =   13170
            TabIndex        =   216
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   91
            Left            =   12885
            TabIndex        =   213
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   90
            Left            =   12885
            TabIndex        =   212
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   89
            Left            =   12600
            TabIndex        =   209
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   88
            Left            =   12600
            TabIndex        =   208
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   87
            Left            =   12315
            TabIndex        =   205
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   86
            Left            =   12315
            TabIndex        =   204
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   85
            Left            =   12030
            TabIndex        =   201
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   84
            Left            =   12030
            TabIndex        =   200
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   83
            Left            =   11745
            TabIndex        =   197
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   82
            Left            =   11745
            TabIndex        =   196
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   81
            Left            =   11460
            TabIndex        =   193
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   80
            Left            =   11460
            TabIndex        =   192
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   79
            Left            =   11175
            TabIndex        =   189
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   78
            Left            =   11175
            TabIndex        =   188
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   77
            Left            =   10890
            TabIndex        =   185
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   76
            Left            =   10890
            TabIndex        =   184
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   75
            Left            =   10605
            TabIndex        =   181
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   74
            Left            =   10605
            TabIndex        =   180
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   73
            Left            =   10320
            TabIndex        =   177
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   72
            Left            =   10320
            TabIndex        =   176
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   71
            Left            =   10035
            TabIndex        =   173
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   70
            Left            =   10035
            TabIndex        =   172
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   69
            Left            =   9750
            TabIndex        =   169
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   68
            Left            =   9750
            TabIndex        =   168
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   67
            Left            =   9465
            TabIndex        =   165
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   66
            Left            =   9465
            TabIndex        =   164
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   65
            Left            =   9180
            TabIndex        =   161
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   64
            Left            =   9180
            TabIndex        =   160
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   63
            Left            =   8895
            TabIndex        =   157
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   62
            Left            =   8895
            TabIndex        =   156
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   61
            Left            =   8610
            TabIndex        =   153
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   60
            Left            =   8610
            TabIndex        =   152
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   59
            Left            =   8325
            TabIndex        =   149
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   58
            Left            =   8325
            TabIndex        =   148
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   57
            Left            =   8040
            TabIndex        =   145
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   56
            Left            =   8040
            TabIndex        =   144
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   55
            Left            =   7755
            TabIndex        =   141
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   54
            Left            =   7755
            TabIndex        =   140
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   53
            Left            =   7470
            TabIndex        =   137
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   52
            Left            =   7470
            TabIndex        =   136
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   51
            Left            =   7185
            TabIndex        =   133
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   50
            Left            =   7185
            TabIndex        =   132
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   49
            Left            =   6900
            TabIndex        =   129
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   48
            Left            =   6900
            TabIndex        =   128
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   47
            Left            =   6615
            TabIndex        =   125
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   46
            Left            =   6615
            TabIndex        =   124
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   45
            Left            =   6330
            TabIndex        =   121
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   44
            Left            =   6330
            TabIndex        =   120
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   43
            Left            =   6045
            TabIndex        =   117
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   42
            Left            =   6045
            TabIndex        =   116
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   41
            Left            =   5760
            TabIndex        =   113
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   40
            Left            =   5760
            TabIndex        =   112
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   39
            Left            =   5475
            TabIndex        =   109
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   38
            Left            =   5475
            TabIndex        =   108
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   37
            Left            =   5190
            TabIndex        =   105
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   36
            Left            =   5190
            TabIndex        =   104
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   35
            Left            =   4905
            TabIndex        =   101
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   34
            Left            =   4905
            TabIndex        =   100
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   33
            Left            =   4620
            TabIndex        =   97
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   32
            Left            =   4620
            TabIndex        =   96
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   31
            Left            =   4335
            TabIndex        =   93
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   30
            Left            =   4335
            TabIndex        =   92
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   29
            Left            =   4050
            TabIndex        =   89
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   28
            Left            =   4050
            TabIndex        =   88
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   27
            Left            =   3765
            TabIndex        =   85
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   26
            Left            =   3765
            TabIndex        =   84
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   25
            Left            =   3480
            TabIndex        =   81
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   24
            Left            =   3480
            TabIndex        =   80
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   23
            Left            =   3195
            TabIndex        =   77
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   22
            Left            =   3195
            TabIndex        =   76
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   21
            Left            =   2910
            TabIndex        =   73
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   20
            Left            =   2910
            TabIndex        =   72
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   19
            Left            =   2625
            TabIndex        =   69
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   18
            Left            =   2625
            TabIndex        =   68
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   17
            Left            =   2340
            TabIndex        =   65
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   16
            Left            =   2340
            TabIndex        =   64
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   15
            Left            =   2055
            TabIndex        =   61
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   14
            Left            =   2055
            TabIndex        =   60
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   13
            Left            =   1770
            TabIndex        =   57
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   12
            Left            =   1770
            TabIndex        =   56
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   11
            Left            =   1485
            TabIndex        =   53
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   10
            Left            =   1485
            TabIndex        =   52
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   9
            Left            =   1200
            TabIndex        =   49
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   8
            Left            =   1200
            TabIndex        =   48
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   7
            Left            =   915
            TabIndex        =   45
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   6
            Left            =   915
            TabIndex        =   44
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   5
            Left            =   630
            TabIndex        =   41
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   4
            Left            =   630
            TabIndex        =   40
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   3
            Left            =   345
            TabIndex        =   37
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   2
            Left            =   345
            TabIndex        =   36
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   33
            Top             =   75
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   1
            Left            =   60
            TabIndex        =   32
            Top             =   510
            Visible         =   0   'False
            Width           =   285
         End
      End
      Begin VB.CommandButton cmdInitPing 
         Caption         =   "Init Ping"
         Enabled         =   0   'False
         Height          =   330
         Left            =   2925
         TabIndex        =   1
         Top             =   60
         Width           =   1080
      End
      Begin VB.CheckBox chkLoadFromFile 
         Caption         =   "Load from file"
         Height          =   255
         Left            =   1485
         TabIndex        =   28
         Top             =   90
         Width           =   1395
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Test1"
         Height          =   330
         Left            =   15345
         TabIndex        =   27
         Top             =   75
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.CheckBox chkAutoSort 
         Caption         =   "Auto Sorting"
         Height          =   255
         Left            =   13575
         TabIndex        =   23
         Top             =   120
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtTest 
         Height          =   285
         Left            =   14430
         TabIndex        =   24
         Top             =   75
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.CommandButton cmdKillAgent 
         Caption         =   "Kill Agent"
         Height          =   330
         Left            =   11985
         TabIndex        =   22
         Top             =   90
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.CommandButton cmdOptions 
         Caption         =   "Options"
         Height          =   330
         Left            =   6165
         TabIndex        =   4
         Top             =   60
         Width           =   1080
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "Stop Ping"
         Enabled         =   0   'False
         Height          =   330
         Left            =   5085
         TabIndex        =   3
         Top             =   60
         Width           =   1080
      End
      Begin VB.CommandButton cmdStart 
         Caption         =   "Ping!"
         Enabled         =   0   'False
         Height          =   330
         Left            =   4005
         TabIndex        =   2
         Top             =   60
         Width           =   1080
      End
      Begin VB.CommandButton cmdLoadList 
         Caption         =   "Load List"
         Height          =   330
         Left            =   45
         TabIndex        =   0
         Top             =   60
         Width           =   1350
      End
   End
   Begin VB.Timer tmrGetReportData 
      Enabled         =   0   'False
      Left            =   8685
      Top             =   3945
   End
   Begin VB.PictureBox picLed 
      Height          =   315
      Index           =   0
      Left            =   7890
      Picture         =   "frmMain.frx":4C3BC
      ScaleHeight     =   255
      ScaleWidth      =   240
      TabIndex        =   20
      Top             =   4605
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picLed 
      Height          =   315
      Index           =   1
      Left            =   8295
      Picture         =   "frmMain.frx":4CDBE
      ScaleHeight     =   255
      ScaleWidth      =   240
      TabIndex        =   19
      Top             =   4605
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picLed 
      Height          =   315
      Index           =   2
      Left            =   8655
      Picture         =   "frmMain.frx":4D7C0
      ScaleHeight     =   255
      ScaleWidth      =   240
      TabIndex        =   18
      Top             =   4605
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picLed 
      Height          =   315
      Index           =   3
      Left            =   9015
      Picture         =   "frmMain.frx":4E1C2
      ScaleHeight     =   255
      ScaleWidth      =   240
      TabIndex        =   17
      Top             =   4605
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Timer tmrCheckAgentAlive 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9390
      Top             =   2655
   End
   Begin MSComctlLib.ImageList imlSmall 
      Left            =   7905
      Top             =   3120
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
            Picture         =   "frmMain.frx":4EBC4
            Key             =   "green"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4F5D6
            Key             =   "red"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4FFE8
            Key             =   "yellow"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":509FA
            Key             =   "gray"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar statusbar 
      Align           =   2  '¹ï»ôªí³æ¤U¤è
      Height          =   330
      Left            =   0
      TabIndex        =   16
      Top             =   5295
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4313
            Text            =   "°T®§"
            TextSave        =   "°T®§"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4313
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4313
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4595
            MinWidth        =   2822
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4313
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4313
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSortIcon 
      Left            =   8670
      Top             =   3090
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
            Picture         =   "frmMain.frx":5140C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":514DE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvPingList 
      Height          =   2610
      Left            =   45
      TabIndex        =   14
      Top             =   2145
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   4604
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      PictureAlignment=   4
      _Version        =   393217
      Icons           =   "imlSmall"
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
      NumItems        =   17
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "#"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Node Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "IP Address"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Description 1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Description 2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Description 3"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Sent "
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "Received"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "Lost"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "Packet Loss%"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   11
         Text            =   "Avg RTT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   12
         Text            =   "Ping Cycle Interval(sec)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   13
         Text            =   "Fail Cycle Count"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   14
         Text            =   "Ping Cycle Count"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Down Event"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Last Down Time"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvDownList 
      Height          =   2610
      Left            =   915
      TabIndex        =   15
      Top             =   2370
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   4604
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      PictureAlignment=   4
      _Version        =   393217
      Icons           =   "imlSmall"
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
      NumItems        =   17
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "#"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Node Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "IP Address"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Description 1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Description 2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Description 3"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   7
         Text            =   "Sent "
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   8
         Text            =   "Received"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "Lost"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   10
         Text            =   "Packet Loss%"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   11
         Text            =   "Avg RTT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   12
         Text            =   "Ping Cycle Interval(sec)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   13
         Text            =   "Fail Cycle Count"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   14
         Text            =   "Ping Cycle Count"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Down Event"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "Last Down Time"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.TabStrip tabx 
      Height          =   4845
      Left            =   45
      TabIndex        =   13
      Top             =   1965
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   8546
      TabFixedWidth   =   2819
      HotTracking     =   -1  'True
      TabMinWidth     =   2118
      ImageList       =   "imlCaption"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Status"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Alert"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private SummaryIsRunning As Boolean
Private LastPingEndNode As Long
Private PingStartNode As Long, PingEndNode As Long

Private PingCount As Long

Private CheckIsRunning As Boolean
Private LastCheckEnd As Long

Private Const WM_SETREDRAW = &HB
Private Const LVM_FIRST As Long = &H1000
Private Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
Private Const LVSCW_AUTOSIZE As Long = -1
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2

Private PingListSortOrder() As Integer
Private DownListSortOrder() As Integer
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long

Private Const WM_QUIT = &H12
Private Const WM_CLOSE = &H10
Private Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hWnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long
Private Type RECT
   Left    As Long
   Top     As Long
   Right   As Long
   Bottom  As Long
End Type
Private Declare Function GetClientRect& Lib "user32" _
                            (ByVal hWnd&, Rct As RECT)

Private SPACE9 As String ''¸ü¤Jfailcount ,successcount¥Î(listview±Æ§Ç)
Private SPACE5 As String
Private SPACE3 As String

Private LogCounter As Integer
Private ChechAgentReadyCount As Long
Private SelectedTab As Integer
Private Const SPACE2 = "  "
Private Sub cmdAbout_Click()
    frmAbout.Show
End Sub

Private Sub cmdExcel_Click()
    Dim excl As Excel.Application
    Dim bk As Workbook
    Dim sht As Worksheet
    Dim i As Integer, j As Integer
    Dim txtcontent As String
    Dim tmpvalue As String
    Dim tmpline As String
    Dim status As String
    Dim lv As ListView
    
    On Error GoTo ErrHandler
    Me.MousePointer = vbHourglass
    If SelectedTab = 1 Then
        Set lv = lvPingList
    Else
        Set lv = lvDownList
    End If
    txtcontent = "Status"
    For i = 2 To lv.ColumnHeaders.Count
        txtcontent = txtcontent & vbTab & lv.ColumnHeaders(i)
    Next i
    
    For i = 1 To lv.ListItems.Count
        status = lv.ListItems(i).Text
        Select Case status
        Case ""
          tmpline = "Y"
        Case " "
          tmpline = "G"
        Case SPACE2
          tmpline = "R"
        End Select
        
        For j = 2 To lv.ColumnHeaders.Count
            tmpvalue = Trim(lv.ListItems(i).SubItems(j - 1))
            Select Case j
            Case 3, 5, 6, 7 'node name, description 1,2,3
                tmpline = tmpline & vbTab & "'" & tmpvalue
            Case Else
                tmpline = tmpline & vbTab & tmpvalue
            End Select
        Next
        txtcontent = txtcontent & vbCr & tmpline
    Next
    Set excl = New Excel.Application
    Set bk = excl.Workbooks.Add
    Set sht = bk.ActiveSheet
    
    sht.Range("A1").Select
    Clipboard.Clear
    Clipboard.SetText txtcontent
    sht.Paste
    
    sht.Columns.AutoFit
    sht.Calculate
    sht.Range("A1").Select
    excl.Visible = True
    Screen.MousePointer = vbDefault
    Me.MousePointer = vbDefault
    Exit Sub
ErrHandler:
    Me.MousePointer = vbDefault
    Exit Sub
End Sub

Private Sub cmdExcelOld_Click()
'http://support.microsoft.com/kb/247412
'Microsoft»¡¥Î³oºØ¤èªk¶Ç¿é¤j¶q¸ê®Æ¨ìExcel·|«ÜºC,´ú¸Õµ²ªG¯uªº¦nºC
    Dim i As Integer, j  As Integer
    Dim excl As Excel.Application
    Dim bk As Workbook
    Dim sht As Worksheet
    Dim cel As Range
    Dim item As ListItem
    Dim subItem As ListSubItem
    Dim status As String
    On Error GoTo ErrHandler
    Screen.MousePointer = vbHourglass
    Set excl = New Excel.Application
    Set bk = excl.Workbooks.Add
    Set sht = bk.ActiveSheet
   
    'Äæ¦ì¼ÐÃD
    For i = 1 To lvPingList.ColumnHeaders.Count
        If i = 1 Then
          sht.Cells(1, i) = "Status"
          sht.Cells(1, i).Font.Bold = True
        Else
          sht.Cells(1, i) = lvPingList.ColumnHeaders(i)
          sht.Cells(1, i).Font.Bold = True
        End If
        
    Next i
    DoEvents
    For i = 1 To lvPingList.ListItems.Count
        status = lvPingList.ListItems(i).Text
        Select Case status
        Case ""
          sht.Cells(i + 1, 1) = "Y"
          sht.Cells(i + 1, 1).Font.Color = RGB(255, 128, 0) 'vbYellow ¥Î¾í¦â,¶À¦â¬Ý¤£²M·¡
        Case " "
          sht.Cells(i + 1, 1) = "G"
          sht.Cells(i + 1, 1).Font.Color = vbGreen 'vbGreen 'green
        Case "  "
          sht.Cells(i + 1, 1) = "R"
          sht.Cells(i + 1, 1).Font.Color = vbRed 'vbRed 'red
        End Select
        sht.Cells(i + 1, 1).Font.Bold = True
        For j = 2 To lvPingList.ColumnHeaders.Count
            sht.Cells(i + 1, j) = lvPingList.ListItems(i).SubItems(j - 1)
        Next j
    Next i

    sht.Columns.AutoFit
    excl.Visible = True
    Screen.MousePointer = vbDefault
    Exit Sub
ErrHandler:
    Screen.MousePointer = vbDefault
    MsgBox "²£¥ÍExcelÀÉ®Éµo¥Í¤F¥H¤Uªº¿ù»~!" & vbCrLf & Err.Description, vbExclamation, MsgTitle
End Sub


Private Sub cmdFindFirst_Click()
    Dim itmX As ListItem
    Dim findstr As String
    Dim i As Integer
    
    Dim lv As ListView
    
    If SelectedTab = 1 Then
        Set lv = lvPingList
    Else
        Set lv = lvDownList
    End If
    
    findstr = Trim(txtFind)
    If findstr = "" Then
        Exit Sub
    End If
    For i = 1 To lv.ListItems.Count
        'Set itmX = lv.ListItems(i).SubItems(2)
        If InStr(1, lv.ListItems(i).SubItems(2), findstr, vbTextCompare) > 0 Then
            Set itmX = lv.ListItems(i)
            itmX.EnsureVisible
            itmX.Selected = True
            lv.SetFocus
            Exit Sub
        End If
    Next
    '¥Î¦¹¤èªk(lvwPartial)¥u¯à·j´MlvwText
    'Set itmX = lv.FindItem(findstr, lvwSubItem, , lvwPartial)
    
    MsgBox "§ä¤£¨ì¦¹¸`ÂI!", vbExclamation, MsgTitle
    txtFind.SelStart = 0
    txtFind.SelLength = Len(txtFind.Text)
    txtFind.SetFocus
        
End Sub

Private Sub cmdFindNext_Click()
    Dim itmX As ListItem
    Dim findstr As String
    Dim findstart As Integer
    Dim i As Integer
    
    Dim lv As ListView
    
    If SelectedTab = 1 Then
        Set lv = lvPingList
    Else
        Set lv = lvDownList
    End If
    
    findstr = Trim(txtFind)
    If findstr = "" Then
        Exit Sub
    End If
    findstart = lv.SelectedItem.Index + 1
    If findstart > lv.ListItems.Count Then
        findstart = 1
    End If
    For i = findstart To lv.ListItems.Count
        'Set itmX = lv.ListItems(i).SubItems(2)
        If InStr(1, lv.ListItems(i).SubItems(2), findstr, vbTextCompare) > 0 Then
            Set itmX = lv.ListItems(i)
            itmX.EnsureVisible
            itmX.Selected = True
            lv.SetFocus
            Exit Sub
        End If
    Next
    '¥Î¦¹¤èªk(lvwPartial)¥u¯à·j´MlvwText
    'Set itmX = lv.FindItem(findstr, lvwSubItem, , lvwPartial)
    MsgBox "§ä¤£¨ì¦¹¸`ÂI!", vbExclamation, MsgTitle
    txtFind.SelStart = 0
    txtFind.SelLength = Len(txtFind.Text)
    txtFind.SetFocus
End Sub

Private Sub cmdGoTo_Click()
    Dim itmX As ListItem
    Dim num As Integer
    Dim i As Integer
    
    Dim lv As ListView
    
    If SelectedTab = 1 Then
        Set lv = lvPingList
    Else
        Set lv = lvDownList
    End If
    
    If IsNumeric(txtGoTo.Text) Then
        num = CInt(txtGoTo.Text)
        
        If num > lv.ListItems.Count Then
            MsgBox "§ä¤£¨ì¦¹¸`ÂI!", vbExclamation, MsgTitle
            txtGoTo.SelStart = 0
            txtGoTo.SelLength = Len(txtGoTo.Text)
            txtGoTo.SetFocus
            num = lv.ListItems.Count
            If num = 0 Then
                Exit Sub
            End If
            
        End If
        
        If num < 1 Then
            num = 1
        End If

        'Set itmX = lv.ListItems(num)
        '§ï¥Î³oºØ¤èªk,¦]¬°itemªº¦ì¸m¥i¯à¦]¬°Äæ¦ìªº±Æ§Ç¦Ó¤£¦P©ósn
        For i = 1 To lv.ListItems.Count
            If lv.ListItems(i).Tag = num Then
                Set itmX = lv.ListItems(i)
                itmX.EnsureVisible
                itmX.Selected = True
                lv.SetFocus
                Exit Sub
            End If
        Next
    Else
        txtGoTo.SelStart = 0
        txtGoTo.SelLength = Len(txtGoTo.Text)
        txtGoTo.SetFocus
    End If
End Sub

Private Sub cmdKillAgent_Click()
    Call TerminateMyAgent
    ReDim aryAgentReady(1 To glAgentCount)

End Sub
Private Sub TerminateMyAgent()
    tmrLog.Enabled = False
    TerminateAppProcess App.Path & "\agent\pingagent.exe"
    
End Sub
Private Sub CloseOldApp()
'·sª©ªº¬yµ{¬O¥ý¸ü¤JPingCollector,µM«áµ¥PingCollector SayHello
    '¥ýÃö±¼ÂÂªºµ{¦¡
    Call TellCollectorToClose
    DoEvents
    
    Call TellAgentToClose
    DoEvents
    
    Dim appcount As Integer
    Dim fs As FileSystemObject
    Dim PingAgentPath As String
    
    
    Set fs = New FileSystemObject
    
    PingAgentPath = App.Path & "\Agent\PingAgent.exe"
    PingAgentPath = fs.GetAbsolutePathName(PingAgentPath)
    If fs.FileExists(PingAgentPath) Then
        appcount = CountAppProcess(PingAgentPath)
        If appcount > 0 Then
            TerminateAppProcess PingAgentPath
            DoEvents
        End If
    End If
    
    Dim PingCollectorPath As String
    
    PingCollectorPath = App.Path & "\agent\pingcollector.exe"
    PingCollectorPath = fs.GetAbsolutePathName(PingCollectorPath)
    
    If fs.FileExists(PingCollectorPath) Then
        appcount = CountAppProcess(PingCollectorPath)
        If appcount > 0 Then
            TerminateAppProcess PingCollectorPath
            DoEvents
        End If
    End If
    
    '§â©Ò¦³ªºtimer¬Ò³]¬°disable
    tmrCheckAgentAlive.Enabled = False
    tmrCheckCollectorAlive.Enabled = False
    tmrChkAgentReady.Enabled = False
    tmrGetReportData.Enabled = False
    
    '***«ì´_ªì©l­È
    Dim i As Integer
    Dim idx As Integer
    If IsArrayInitialized(aryAgentHwnd) Then
        For i = 1 To glAgentCount
            aryAgentHwnd(i) = 0
            aryAgentReady(i) = False
        Next
    End If
    For idx = 0 To glAgentCount - 1
        picAgent(idx).Picture = picLed(0).Picture '¦Ç¦â
        picAgent(idx).Refresh
    Next
    picPingCollector.Picture = picLed(0).Picture '¦Ç¦â
    picPingCollector.Refresh
    glCollectorReady = False
    glPingCollectorHwnd = 0
    
End Sub
Private Sub cmdInitPing_Click()
    If NumOfPingNode = 0 Then
        MsgBox "Please run LoadList first!", vbExclamation, MsgTitle
        Exit Sub
    End If
        
    If glAgentCount = 0 Then
        MsgBox "PingAgent count can not be zeor. Please change the options!", vbExclamation, MsgTitle
        Exit Sub
    End If
    If glAgentCount > MaxAgent Then
        MsgBox "PingAgent count can not exceed " & MaxAgent & "! Please change the options!", vbExclamation, MsgTitle
        Exit Sub
    End If
    
    If NumOfPingNode < glAgentCount Then
        MsgBox "PingAgent count is more than ping nodes! Please change the options!", vbExclamation, MsgTitle
        Exit Sub
    End If
    
    '***¥i¥H¶}©l¤F
    Call CloseOldApp
    
    Call InitAgentLedToGrayColor '¥þ³¡led³]¦¨¦Ç¦â
    Call ResizeForm
    picPingCollector.Picture = picLed(0).Picture '¥ý³]¦¨¦Ç¦â
    picPingCollector.Refresh
    
    '¦bcreate agent¤§«eÅª¨ú°Ñ¼Æ,¨Ã¤@¨Ö¶Çµ¹agent
    'start agent --> agent say hello --> tell agent to load ini --> agent report load ini ok -->
    'tell agent to load ip list --> agent report ready!
    
    aryINI(1) = glPingCount
    aryINI(2) = glPingTimeOutHost
    aryINI(3) = glPingMaxBurst
    aryINI(4) = glPingTimeOutBatch
    aryINI(5) = glWaitForSingleObject
    aryINI(6) = glCycleInterval
    
    aryINI(7) = glAgentCount
    aryINI(8) = glThreshold
    aryINI(9) = glContinuedFailAsDown
    aryINI(10) = glStatisticsCycle
    aryINI(11) = glRefreshCycle
    aryINI(12) = glDelayStart
    
    Dim LoadOK As Boolean
    '½ÐPingCollector¥ý¸ü¤JINI (¨Ãªì©l¤ÆAgent Array)
    
    LoadOK = CreatePingCollector '·|¶Ç¦^°T®§,¬G¥²¶·­n¥ý°õ¦æHook
    If LoadOK Then
        '-->Collector SayHello --> TellCollectorToLoadList --> LoadList OK --> TellCollectorToLoadINI -->
        '-->Collector Ready --> CreatePingAgent -->
    End If




''    '¦]¬°pingagent°_¨Ó«á·|¦Vcollector say hello!, ¦Ócollector¥²¶·¥ýÅª¶iini­È·í¤¤ªº­È,¨Ò¦pglPingCount,¤~¯àªì©l¤Æ
''    '°}¦C, ¦b¦¬¨ì
''    Call TellCollectorToLoadINI ' --> PingCollectorª½±µ¦^³øReady--> PingCollectorReady --> createpingagent
    
End Sub

Private Function CreatePingAgent() As Boolean
    Dim AgentID As Integer
    Dim ret As Long
    Dim fs As FileSystemObject
    Dim PingAgentPath As String
    
    
    Set fs = New FileSystemObject
    PingAgentPath = App.Path & "\Agent\PingAgent.exe"
    PingAgentPath = fs.GetAbsolutePathName(PingAgentPath)
    If Not fs.FileExists(PingAgentPath) Then
        MsgBox "§ä¤£¨ìPingAgentµ{¦¡! (" & PingAgentPath & ")", vbExclamation, MsgTitle
        CreatePingAgent = False
        Exit Function
    End If
    
    ReDim aryAgentHwnd(1 To glAgentCount)
    ReDim aryAgentReady(1 To glAgentCount)
    ReDim aryAgentBlinkTick(1 To glAgentCount)
    
    For AgentID = 1 To glAgentCount
        'command line argument: me.hwnd agentid
        ret = Shell(PingAgentPath & " " & Me.hWnd & Space(1) & _
            AgentID & Space(1) & _
            glPingCollectorHwnd & Space(1) & _
            IIf(glDebugMode, 1, 0), vbNormalNoFocus)
        If ret = 0 Then
            MsgBox "°õ¦æCreatePingAgent®Éµo¥Í¿ù»~!", vbExclamation, MsgTitle
            CreatePingAgent = False
            Exit Function
        End If
    Next
       
    CreatePingAgent = True
End Function
'Private Function LoadLogAgent() As Boolean

Private Function CreatePingCollector() As Boolean
    'log agent¤£¯à¦bForm Load®É¸ü¤J,§_«h¦¹¥Dµ{¦¡©|¥¼§@¥Î,¦¬¤£¨ì¥Ñlog agentµo¨Óªº°T®§
    Dim ret As Long
    Dim fs As FileSystemObject
    Dim PingCollectorPath As String
    Dim appcount As Integer
    
    Set fs = New FileSystemObject
    
    On Error GoTo ErrHandler
    
    CreatePingCollector = False
    
    PingCollectorPath = App.Path & "\agent\pingcollector.exe"
    PingCollectorPath = fs.GetAbsolutePathName(PingCollectorPath)
    
    If Not fs.FileExists(PingCollectorPath) Then
        MsgBox "§ä¤£¨ìPingCollectorµ{¦¡! (" & PingCollectorPath & ")", vbExclamation, MsgTitle
        Exit Function
    End If
    
    glPingCollectorHwnd = 0
    
    If glDebugMode Then
        'MsgBox glDebugMode
        ret = Shell(PingCollectorPath & " " & Me.hWnd & " 1", vbMinimizedFocus)
    Else
        ret = Shell(PingCollectorPath & " " & Me.hWnd & " 0", vbHide)
    End If
    If ret = 0 Then
        MsgBox "°õ¦æCreatePingCollector®Éµo¥Í¿ù»~!", vbExclamation, MsgTitle
        Exit Function
    End If
    CreatePingCollector = True
    Exit Function
ErrHandler:
    
End Function

Private Sub InitPingListSort()
    Dim col As ColumnHeader
    Set col = lvPingList.ColumnHeaders(2)     'sort on first column
    lvPingList.ColumnHeaderIcons = imlSortIcon
    PingListSortOrder(2) = lvwDescending      'will get flipped to ascending
    lvPingList_ColumnClick col                'click the column heading
    If lvPingList.ListItems.Count > 0 Then
        lvPingList.ListItems(1).EnsureVisible     'make sure first one is visible
        lvPingList.ListItems(1).Selected = True   'and selected
    End If
End Sub
Private Sub InitDownListSort()
    Dim col As ColumnHeader
    Set col = lvDownList.ColumnHeaders(2)     'sort on 2th column: description 1
    lvDownList.ColumnHeaderIcons = imlSortIcon
    DownListSortOrder(2) = lvwDescending      'will get flipped to ascending
    lvDownList_ColumnClick col                'click the column heading
    If lvDownList.ListItems.Count > 0 Then
        lvDownList.ListItems(1).EnsureVisible     'make sure first one is visible
        lvDownList.ListItems(1).Selected = True   'and selected
    End If
End Sub

Private Sub cmdLoadList_Click()
    Dim LoadOK As Boolean
    
    cmdInitPing.Enabled = False
    
    If chkLoadFromFile.Value = 1 Then
        LoadOK = LoadListFromFile
        If Not LoadOK Then
            Exit Sub
        End If
        frmSavePingListIntoDB.Show vbModal, Me
        chkLoadFromFile.Value = 0
        If Not LoadListIntoDBOK Then
            Exit Sub
        End If
        'frmDupCheck.Show vbModal, Me
        If NumOfPingNode > 1 Then
            frmChkDupNode.Show vbModal, Me
            If Not CheckDupNodeOK Then
                Exit Sub
            End If
        End If
    End If
    '¦pªG±qÀÉ®×¸ü¤J --> DB --> Load From DB
    LoadOK = LoadListFromDB
    If Not LoadOK Then
        Exit Sub
    End If
    If NumOfPingNode > 0 Then
        '***¥i¥H¶}©l¤F
        Call CloseOldApp
        Call InitAgentLedToGrayColor
        Call ResizeForm
    
        ReDim aryReportData(1 To 8, MaxNodeIndex)
        ReDim aryLastLedResult(MaxNodeIndex)

        LoadPingList
        
        statusbar.Panels(2).Text = ""
        statusbar.Panels(3).Text = ""
        statusbar.Panels(4).Text = ""
        statusbar.Panels(5).Text = ""
        ReDim PingListSortOrder(1 To lvPingList.ColumnHeaders.Count)
        ReDim DownListSortOrder(1 To lvDownList.ColumnHeaders.Count)
        Call InitPingListSort
        
        cmdStart.Enabled = False
        cmdLoadList.Enabled = True
        cmdInitPing.Enabled = True
        cmdStop.Enabled = False
    End If
End Sub
Public Sub PingCollectorLoadPingListOK()
     ShowStatus "Collector Load INI..."
     Call TellCollectorToLoadINI
End Sub
Public Sub PingCollectorLoadPingListErr()
    picPingCollector.Picture = picLed(3).Picture '¬õ¦â
    picPingCollector.Refresh
End Sub
Private Function LoadListFromFile() As Boolean
    Dim fs As FileSystemObject
    Dim F As TextStream
    Dim tmpline As String
    Dim aryCol() As String
    Dim i As Integer
    
    On Error GoTo ErrHandler
    LoadListFromFile = False
    '¥ý¸ü¤JÀÉ®×¨ìarray
    Set fs = New FileSystemObject
    glPingListFile = App.Path & "\pinglist.txt"
    If Not fs.FileExists(glPingListFile) Then
        MsgBox glPingListFile & " ÀÉ®×¤£¦s¦b!", vbExclamation, MsgTitle
        Exit Function
    End If
    Set F = fs.OpenTextFile(glPingListFile, ForReading, False)
    
    ReDim aryNodeName(MAX_PING_NODES)
    ReDim aryIPAddress(MAX_PING_NODES)
    ReDim aryInetAddr(MAX_PING_NODES)
    ReDim aryRoute1(MAX_PING_NODES)
    ReDim aryRoute2(MAX_PING_NODES)
    ReDim aryRoute3(MAX_PING_NODES)
    ReDim arySN(MAX_PING_NODES)
    
    i = 0
    
    Do While F.AtEndOfStream <> True
        tmpline = Trim(F.ReadLine)
        If Left(tmpline, 1) <> "#" And Len(tmpline) <> 0 Then
            aryCol = Split(tmpline, ",")
            If UBound(aryCol) >= 4 Then
                aryNodeName(i) = Trim(aryCol(0))
                aryIPAddress(i) = Trim(aryCol(1))
                aryInetAddr(i) = inet_addr(aryIPAddress(i))
                aryRoute1(i) = Trim(aryCol(2))
                aryRoute2(i) = Trim(aryCol(3))
                aryRoute3(i) = Trim(aryCol(4))
                arySN(i) = i + 1
                i = i + 1
            ElseIf UBound(aryCol) >= 3 Then
                aryNodeName(i) = Trim(aryCol(0))
                aryIPAddress(i) = Trim(aryCol(1))
                aryInetAddr(i) = inet_addr(aryIPAddress(i))
                aryRoute1(i) = Trim(aryCol(2))
                aryRoute2(i) = Trim(aryCol(3))
                arySN(i) = i + 1
                i = i + 1
            ElseIf UBound(aryCol) >= 2 Then
                aryNodeName(i) = Trim(aryCol(0))
                aryIPAddress(i) = Trim(aryCol(1))
                aryInetAddr(i) = inet_addr(aryIPAddress(i))
                aryRoute1(i) = Trim(aryCol(2))
                arySN(i) = i + 1
                i = i + 1
            ElseIf UBound(aryCol) = 1 Then
                aryNodeName(i) = aryCol(0)
                aryIPAddress(i) = Trim(aryCol(1))
                aryInetAddr(i) = inet_addr(aryIPAddress(i))
                arySN(i) = i + 1
                i = i + 1
            End If
        End If
    Loop
    
    NumOfPingNode = i
    MaxNodeIndex = i - 1
    
    If NumOfPingNode = 0 Then
        Erase aryNodeName
        Erase aryIPAddress
        Erase aryInetAddr
        Erase aryRoute1
        Erase aryRoute2
        Erase aryRoute3
        Erase arySN
        Exit Function
    End If
    ReDim Preserve aryNodeName(MaxNodeIndex)
    ReDim Preserve aryIPAddress(MaxNodeIndex)
    ReDim Preserve aryInetAddr(MaxNodeIndex)
    ReDim Preserve aryRoute1(MaxNodeIndex)
    ReDim Preserve aryRoute2(MaxNodeIndex)
    ReDim Preserve aryRoute2(MaxNodeIndex)
    ReDim Preserve arySN(MaxNodeIndex)
    LoadListFromFile = True
    Exit Function
ErrHandler:
    MsgBox "¥Ñpinglist.txt¸ü¤J¸ê®Æ®Éµo¥Í¿ù»~:" & vbCrLf & Err.Description, vbExclamation, MsgTitle
End Function
Private Function LoadListFromDB() As Boolean
    Dim rsPingList As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo ErrHandler
    LoadListFromDB = False
    Set rsPingList = New ADODB.Recordset
    With rsPingList
        .CursorLocation = adUseClient
        .Open "select * from PingList order by SN;", ConnStr, adOpenDynamic, adLockOptimistic
        If .RecordCount <= 0 Then
            .Close
            Set rsPingList = Nothing
            MsgBox "¸ê®Æ®wpinglist¨S¦³¥ô¦ó¸ê®Æ!", vbExclamation, MsgTitle
            Exit Function
        End If
        
        NumOfPingNode = .RecordCount
        MaxNodeIndex = NumOfPingNode - 1
        ReDim aryNodeName(MaxNodeIndex)
        ReDim aryIPAddress(MaxNodeIndex)
        ReDim aryInetAddr(MaxNodeIndex)
        ReDim aryRoute1(MaxNodeIndex)
        ReDim aryRoute2(MaxNodeIndex)
        ReDim aryRoute3(MaxNodeIndex)
        ReDim arySN(MaxNodeIndex)
        .MoveFirst
        i = 0
        While Not .EOF
            arySN(i) = !SN
            aryNodeName(i) = !NodeName
            aryRoute1(i) = !Route1
            aryRoute2(i) = !Route2
            aryRoute3(i) = !Route3
            aryIPAddress(i) = !IP
            aryInetAddr(i) = inet_addr(aryIPAddress(i))
            i = i + 1
            .MoveNext
        Wend
        .Close
        Set rsPingList = Nothing
    End With
    LoadListFromDB = True
    Exit Function
ErrHandler:
    MsgBox "¥ÑDB¸ü¤Jping list®Éµo¥Í¿ù»~:" & vbCrLf & Err.Description, vbExclamation, MsgTitle
End Function
'Private Sub cmdOpenLog_Click()
'    Dim ret
'    On Error GoTo ErrHandler
'    ret = Shell("notepad.exe " & GetLogFileName, vbMaximizedFocus)
'    If ret <> 0 Then
'        AppActivate ret
'    End If
'    Exit Sub
'ErrHandler:
'    MsgBox "¶}±ÒlogÀÉ®Éµo¥Í¥H¤U¿ù»~:" & vbCrLf & Err.Description, vbExclamation, MsgTitle
'End Sub

Private Sub cmdOpenEventLog_Click()
    Dim frm As frmLog
    On Error GoTo ErrHandler
    
    Set frm = New frmLog '¦¹®É¤w°õ¦æform Load event
    frm.Caption = "log - All Nodes"
    frm.SetNodeName "All Nodes"
    frm.SetIP "*.*.*.*"
    frm.SetDesc "All Alert Events"
    frm.SetShowNodeName True
    frm.LoadLogData
    frm.Show vbModeless
    Exit Sub
ErrHandler:
    MsgBox "Error!" & vbCrLf & Err.Description, vbExclamation, MsgTitle
End Sub

Private Sub cmdOptions_Click()
    frmOptions.Show vbModal, Me
End Sub



Private Sub cmdStart_Click()
    Dim i As Long
    Dim ReadyAgentCount As Integer
    
    If NumOfPingNode = 0 Then
        MsgBox "Please load data first!", vbExclamation, MsgTitle
        Exit Sub
    End If
    
    ReadyAgentCount = 0
    For i = 1 To glAgentCount
        If aryAgentReady(i) Then
            ReadyAgentCount = ReadyAgentCount + 1
        End If
    Next
    
  
    Call ResetEventBuf
    
    UserStop = False
    statusbar.Panels(2).Text = ""
    statusbar.Panels(3).Text = ""
    statusbar.Panels(4).Text = ""
    statusbar.Panels(5).Text = ""
    
    TellAppToDoSomething MSG_PINGCOLLECTOR_STARTPING, glPingCollectorHwnd
    PingIsRunning = True
    tmrGetReportData.Interval = glRefreshCycle
    tmrGetReportData.Enabled = True
    Call StartLog
    cmdStart.Enabled = False
    cmdLoadList.Enabled = False
    cmdInitPing.Enabled = False
    cmdStop.Enabled = True
    '­«·s³]©wtick­È,¦b±Ò°Êcheckagent & checkcollector«e
    'ReDim aryAgentBlinkTick(1 To glAgentCount)

    
    
''    tmrCheckAgent.Enabled = True
''    '³°Äò±Ò°Ê¨C­Óagent,¦Ó¤£¤@¦¸±Ò°Ê

''    tmrLog.Enabled = True


''    tmrGetReportData.Interval = glRefreshCycle
''    tmrGetReportData.Enabled = True
    
    
End Sub
Private Sub ResetEventBuf()
    Erase aryEventBuf1
    Erase aryEventBuf2
    EventBufCount = 0
End Sub
Private Sub cmdStop_Click()
    Call StopPing
    cmdStop.Enabled = False
    cmdLoadList.Enabled = True
End Sub
Public Sub StopPing()

    'tmrSummary.Enabled = False
    
    tmrGetReportData.Enabled = False
    tmrLog.Enabled = False
    
    Dim i As Integer
    If IsArrayInitialized(aryAgentHwnd) Then
        For i = 1 To glAgentCount
            If aryAgentHwnd(i) <> 0 Then
                TellAppToDoSomething MSG_STOP_PING, aryAgentHwnd(i)
            End If
        Next
        DoEvents
    End If
 
    If glPingCollectorHwnd <> 0 Then
        TellAppToDoSomething MSG_STOP_PING, glPingCollectorHwnd
        DoEvents
    End If
    
    PingIsRunning = False

End Sub

Private Sub Command1_Click()
    TellAppToDoSomething MSG_PLS_PINGCOLLECTOR_REPORT_PING_RESULT, glPingCollectorHwnd
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandler
    Me.Caption = MsgTitle
    glMyHwnd = Me.hWnd
    
    glMyPingMsg = RegisterWindowMessage(MSG_MYPING_POSTMSG)
    If glMyPingMsg = 0 Then
        MsgBox "Windows µLªkµù¥UMSG_MYPING_POSTMSG! µ{¦¡³q°T±N¨ü¨ì¼vÅT!", vbExclamation, MsgTitle
    End If
    
    cmdStart.Enabled = False
    cmdLoadList.Enabled = True
    cmdInitPing.Enabled = False
    cmdStop.Enabled = False
        
    ReDim PingListSortOrder(1 To lvPingList.ColumnHeaders.Count)
    ReDim DownListSortOrder(1 To lvDownList.ColumnHeaders.Count)
    
    tabx.Left = 45
    lvPingList.Left = 45
    
    Dim i As Integer
    For i = 0 To MaxAgent - 1
        lblAgent(i).Caption = i + 1
    Next
    Call GetIniInfo
    Call InitAgentLedToGrayColor
    
    'txtTest.Text = Me.hwnd
'    '¥ý«Å§i,§_«hHook·|¾É­Pµ{¦¡µ²§ô(­ì¨Ó¤£·|)
''    ReDim aryAgentLB(1 To glAgentCount)
''    ReDim aryAgentUB(1 To glAgentCount)
''
    '
    SPACE9 = Space(9)
    SPACE5 = Space(5)
    SPACE3 = Space(3) '¥H¤W¬Ò¬O¬°¤Flistview±Æ§Ç¥Î,¤£¬O¬°¤FÅã¥Ü,¦]¬°Åã¥Ü¤W°£¤F©T©w¤j¤p¦rÅé¥~,·|¦]¦r«¬ªº½t¬G,¦r¤¸·|¦³¤j¤p¤§¤À
    
    'ReDim aryAgentReady(1 To glAgentCount) '¥ý«Å§i,¥H§K¦bPing!ÀË¬d®Éµo¥Í¿ù»~!
    
    'ªì©l¤Ætab
    SelectedTab = 1
    tabx.SelectedItem = tabx.Tabs(SelectedTab)
    
    'log path
    
    Dim fs As New FileSystemObject
    Dim txtfile As TextStream
    Dim logfolder As String
    
    logfolder = App.Path & "\log"
    If Not fs.FolderExists(logfolder) Then
      fs.CreateFolder logfolder
    End If
    
    Dim PingLogDBFile As String
    PingLogDBFile = App.Path & "\pinglog.mdb"
    If Not fs.FileExists(PingLogDBFile) Then
        MsgBox PingLogDBFile & " ¸ê®Æ®wÀÉ®×¤£¦s¦b!", vbCritical, MsgTitle
        End
    End If
   
    ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & PingLogDBFile & ";Persist Security Info=False"
    Call Hook
    
    
    Exit Sub
ErrHandler:
    MsgBox Err.Description, vbCritical, MsgTitle
    End
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Dim frm    As Form
    
    On Error Resume Next
    Dim response, msg, Style, title

    title = MsgTitle  ' ©w¸q¼ÐÃD
    
    Style = vbYesNo + vbCritical + vbDefaultButton1   ' ©w¸q«ö¶s
    msg = "Are you sure you want to quit?"   ' ©w¸q°T®§
    response = MsgBox(msg, Style, title)
    If response = vbNo Then   ' ­Y¨Ï¥ÎªÌ«ö¤U [§_]
        Cancel = True
        Exit Sub
    End If
    
    
'    For Each frm In Forms
'        Unload frm
'    Next
    
    On Error Resume Next
    tmrCheckAgentAlive.Enabled = False
    tmrLog.Enabled = False
    tmrCheckCollectorAlive.Enabled = False
    tmrGetReportData.Enabled = False
    tmrChkAgentReady.Enabled = False
    UnHook '¥ýunhook, ³o¼Ë´N¤£·|±µ¦¬¨Ó¦Üagent & collectorªº°T®§¤F
    Call TellAgentToClose
    DoEvents
    Call TellCollectorToClose
    DoEvents
    Dim frm    As Form
    For Each frm In Forms
        Unload frm
    Next
   
End Sub
Public Sub ResizeForm()
    Call Form_Resize
End Sub
Private Sub Form_Resize()
    Dim clientrect As RECT
    Dim w As Long
    On Error GoTo ErrHandler
    If Me.WindowState = vbMinimized Then Exit Sub
    GetClientRect Me.hWnd, clientrect
    Picture1.Width = (clientrect.Right - clientrect.Left) * Screen.TwipsPerPixelX - 120
    'Bevel1.Width = Picture1.Width - 120
'    If frLedPanel.Width > picLeds.Width Then
'        picLeds.Left = (frLedPanel.Width - picLeds.Width) / 2
'    Else
'        picLeds.Left = 45
'    End If
    'ª`·N:¤£­n¥Îme.Height,¤@¥¹¤Á´«classic style / xp style¤£¦Pªºcontrol®É,¶ZÂ÷·|¦³»~®t
    
    w = (clientrect.Right - clientrect.Left) * Screen.TwipsPerPixelX
    
    Picture1.Left = 60
    Picture1.Width = w - 120
    Call PosAgentLed
    tabx.Top = picToolbar.Height + 10
    tabx.Height = (clientrect.Bottom - clientrect.Top) * Screen.TwipsPerPixelY - picToolbar.Height - statusbar.Height - 25
    tabx.Width = w - 90
    
    lvPingList.Top = tabx.ClientTop + 10
    lvPingList.Left = tabx.ClientLeft '+ 10
    lvPingList.Width = tabx.ClientWidth - 30
    lvPingList.Height = tabx.ClientHeight - 20
    
    lvDownList.Top = tabx.ClientTop + 10
    lvDownList.Left = tabx.ClientLeft '+ 10
    lvDownList.Width = tabx.ClientWidth - 30
    lvDownList.Height = tabx.ClientHeight - 20
    
    AdjustColWidth lvPingList
    AdjustColWidth lvDownList
    Exit Sub
ErrHandler:

End Sub
Private Sub PosAgentLed()
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim xgap As Integer
    Dim xnum As Integer
    Dim totalnum As Integer
    Dim colcount As Integer
    Const ynum As Integer = 2 '³Ì¦hÅã¥Ü2¦C
    

'    For i = 0 To MaxAgent - 1
'        lblAgent(i).Visible = False
'        picAgent(i).Visible = False
'    Next
    
    xnum = (Picture1.Width - 150) / 285 '¨C­Óledªº¶¡¹j¬°285
    If xnum > 0 Then
        k = 0
        If xnum < glAgentCount Then
            colcount = 2
            Picture1.Height = 990
            Picture2.Top = 1550
            picToolbar.Height = 1895
            
            totalnum = xnum * ynum
            If totalnum > glAgentCount Then
                totalnum = glAgentCount
            End If
            If (totalnum Mod 2) = 1 Then
                xnum = totalnum / 2 + 1
            Else
                xnum = totalnum / 2
            End If
            
            For i = 0 To xnum - 1
                lblAgent(k).Top = 75
                picAgent(k).Top = 270
                xgap = 60 + i * 285
                lblAgent(k).Left = xgap
                picAgent(k).Left = xgap
                lblAgent(k).Visible = True
                picAgent(k).Visible = True
                k = k + 1
                If k >= totalnum Then
                    Exit For
                End If
                
                lblAgent(k).Top = 510
                picAgent(k).Top = 705
                xgap = 60 + i * 285
                lblAgent(k).Left = xgap
                picAgent(k).Left = xgap
                lblAgent(k).Visible = True
                picAgent(k).Visible = True
                k = k + 1
                If k >= totalnum Then
                    Exit For
                End If
            Next
        Else
            colcount = 1 '¤@¦C§Y¥i
            Picture1.Height = 495
            Picture2.Top = 1055
            picToolbar.Height = 1400
            totalnum = glAgentCount
            
            For i = 0 To xnum - 1
                lblAgent(k).Top = 75
                picAgent(k).Top = 270
                xgap = 60 + i * 285
                lblAgent(k).Left = xgap
                picAgent(k).Left = xgap
                lblAgent(k).Visible = True
                picAgent(k).Visible = True
                k = k + 1
                If k >= totalnum Then
                    Exit For
                End If
            Next
        End If

        For i = k To MaxAgent - 1
            lblAgent(i).Visible = False
            picAgent(i).Visible = False
        Next
    End If
End Sub
Private Sub LoadPingList(Optional IsInit As Boolean = False)
    Dim x As Long
    Dim i As Long
    Dim j As Long
    Dim itemx As ListItem
    Dim SN As String
    Dim sbuffer As String
    Dim maxnumlen As Integer
    On Error GoTo ErrHandler
    x = SendMessage(lvPingList.hWnd, WM_SETREDRAW, 0, 0)
    maxnumlen = Len(CStr(NumOfPingNode))
    sbuffer = Space(maxnumlen)
    
    With lvPingList
'        Set .SmallIcons = imlSmall
'        '.SmallIcons = imlSmall
        .ListItems.Clear
        
        '¸ü¤J¸ê®Æ
        If Not IsInit Then
            lvPingList.ColumnHeaders(1).Text = ""
            lvPingList.ColumnHeaders(1).Width = 300
            For i = 0 To NumOfPingNode - 1
                'Add(index, key, text, icon, smallIcon)
                SN = Right(sbuffer & arySN(i), maxnumlen)
                'Set itemx = .ListItems.Add(, "#" & i, "", "gray", "gray")
                Set itemx = .ListItems.Add(, "#" & i, " ", , "green")
                itemx.Tag = arySN(i) '°O¿ýi,¤ñkey¦n¥Î
                itemx.SubItems(1) = SN
                itemx.SubItems(2) = aryNodeName(i)
                itemx.SubItems(3) = aryIPAddress(i)
                itemx.SubItems(4) = aryRoute1(i)
                itemx.SubItems(5) = aryRoute2(i)
                itemx.SubItems(6) = aryRoute3(i)
            Next
        End If

'        ColorGrid lvPingList
    End With
    x = SendMessage(lvPingList.hWnd, WM_SETREDRAW, 1, 0)
    AdjustColWidth lvPingList '­n©ñ¦bWM_SETREDRAW¤§«á,§_«h¹Ï¥Ü·|¦³´Ý¼v
    statusbar.Panels(1).Text = "Á`¦@ " & NumOfPingNode & " ¸`ÂI"
    Exit Sub

ErrHandler:
        'Screen.MousePointer = vbDefault
        MsgBox "¸ü¤JPing List®Éµo¥Í¿ù»~!" & vbCrLf & Err.Description, vbCritical, MsgTitle
        Exit Sub
End Sub

Private Sub AdjustColWidth(lv As ListView)
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
            'If col2adjust <> 1 Then 'ÁôÂÃÄæ¦ì
                Call SendMessage(lv.hWnd, _
                     LVM_SETCOLUMNWIDTH, _
                     col2adjust, _
                     ByVal LVSCW_AUTOSIZE_USEHEADER)
            'End If
        Next
        
    End If
End Sub





Private Sub lvDownList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

' °õ¦æ±Æ§Ç
' 1 is the first column
    Dim i As Integer
    With lvDownList
        If lvDownList.ListItems.Count = 0 Then Exit Sub
        For i = 1 To .ColumnHeaders.Count
            If i = ColumnHeader.Index Then
                DownListSortOrder(i) = FlipSort(DownListSortOrder(ColumnHeader.Index))
            Else
                DownListSortOrder(i) = lvwDescending '¨ä¥¦ªº³]¬°Descending,¤U¤@¦¸click®ÉÅÜ¬°Ascending
            End If
        Next
        .SortOrder = DownListSortOrder(ColumnHeader.Index)
        .SortKey = ColumnHeader.Index - 1
        .Sorted = True
        DoEvents
        
            .Sorted = False
        'Show column icon
        ShowListViewSortIcon lvDownList
        
'        If Not .SelectedItem Is Nothing Then
'            .SelectedItem.EnsureVisible
'        End If
        .ListItems(1).EnsureVisible     'make sure first one is visible
        .ListItems(1).Selected = True   'and selected
        
    End With
    

End Sub

Private Sub lvPingList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
' °õ¦æ±Æ§Ç
' 1 is the first column
    Dim i As Integer
    With lvPingList
        If lvPingList.ListItems.Count = 0 Then Exit Sub
        For i = 1 To .ColumnHeaders.Count
            If i = ColumnHeader.Index Then
                PingListSortOrder(i) = FlipSort(PingListSortOrder(ColumnHeader.Index))
            Else
                PingListSortOrder(i) = lvwDescending '¨ä¥¦ªº³]¬°Descending,¤U¤@¦¸click®ÉÅÜ¬°Ascending
            End If
        Next
        .SortOrder = PingListSortOrder(ColumnHeader.Index)
        .SortKey = ColumnHeader.Index - 1
        .Sorted = True
        DoEvents
        If chkAutoSort.Value <> 1 Then
            .Sorted = False
        End If
        'Show column icon
        ShowListViewSortIcon lvPingList
        
'        If Not .SelectedItem Is Nothing Then
'            .SelectedItem.EnsureVisible
'        End If
        .ListItems(1).EnsureVisible     'make sure first one is visible
        .ListItems(1).Selected = True   'and selected
        
    End With
    
End Sub


Private Sub lvPingList_DblClick()
    Dim itemx As ListItem
    Dim NodeName As String
    Dim frm As frmLog
    
    On Error GoTo ErrHandler
    
    If lvPingList.SelectedItem Is Nothing Then
        
         Exit Sub
    End If
    Set itemx = lvPingList.SelectedItem '¥²¶·©MMouseDown event¦P®É§@·~
    
    NodeName = itemx.SubItems(2)
    If NodeName = "" Then Exit Sub

    Set frm = New frmLog '¦¹®É¤w°õ¦æform Load event
    frm.Caption = "log - " & NodeName
    frm.SetNodeName NodeName
    frm.SetIP itemx.SubItems(3)
    frm.SetDesc itemx.SubItems(4), itemx.SubItems(5), itemx.SubItems(6)
    frm.SetShowNodeName False
    frm.LoadLogData NodeName
    frm.Show vbModeless
    Exit Sub
ErrHandler:
    MsgBox "Error!" & vbCrLf & Err.Description, vbExclamation, MsgTitle
End Sub

Private Sub lvPingList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Set lvPingList.SelectedItem = lvPingList.HitTest(x, y)
End Sub

Private Sub tabx_Click()
    SelectedTab = tabx.SelectedItem.Index
    If SelectedTab = 1 Then
        lvPingList.Visible = True
        lvDownList.Visible = False
    Else
        lvPingList.Visible = False
        lvDownList.Visible = True
        Call ShowDownList
    End If
End Sub
Private Sub ShowDownList()
    Dim itemx As ListItem
    Dim i As Integer
    Dim j As Integer
    Dim colmax As Integer
    Dim x As Long
    Dim idx As Integer
    
    If lvPingList.ListItems.Count = 0 Then Exit Sub
    
    x = SendMessage(lvDownList.hWnd, WM_SETREDRAW, 0, 0)
    
    lvDownList.ListItems.Clear
    lvDownList.ColumnHeaders(1).Width = 300
    colmax = lvPingList.ColumnHeaders.Count
    
    For i = 1 To lvPingList.ListItems.Count
        
        'If aryLastLedResult(i) = RESULT_DOWN Then '¿ù»~ªº¤èªk,¦]¬°listview¥i¯à¸g¹L±Æ§Ç,©M­ì¨Óªºlist order¤£¤@¼Ë¤F
        If lvPingList.ListItems(i).Text = SPACE2 Then
            Set itemx = lvDownList.ListItems.Add(, , SPACE2, , "red")
            For j = 1 To colmax - 1
                itemx.SubItems(j) = lvPingList.ListItems(i).SubItems(j)
            Next
        End If
    Next
        
    x = SendMessage(lvDownList.hWnd, WM_SETREDRAW, 1, 0)
    If lvDownList.ListItems.Count > 0 Then
        AdjustColWidth lvDownList
        InitDownListSort
    End If
End Sub
Private Sub tmrChkAgentReady_Timer()
    'ÀË¬dping agent¬O§_ready
    '   pingagent report ready --> aryAgentReady(agentid) = true --> all agent id report ready then ready
    '¦]¬°¦pªGª½±µ¦bpingagent keep alive reply±NAgentReadyCount + 1, ·|¥Ñ©óagent¦P®É¦^À³ªºÃö«Y,³y¦¨count¥¢¯u
    Dim i As Integer
    Dim AgentReadyCount As Integer
    AgentReadyCount = 0
    For i = 1 To glAgentCount
        If aryAgentReady(i) Then
            AgentReadyCount = AgentReadyCount + 1
        End If
    Next
    '¥þ³¡ªºAgent¬Ò¤wReport Ready,§Ú¤]Report Ready
    If AgentReadyCount = glAgentCount Then
        tmrChkAgentReady.Enabled = False
        cmdStart.Enabled = True
        tmrCheckAgentAlive.Interval = 2000
        tmrCheckAgentAlive.Enabled = True
        ShowStatus "Ready!"
    End If
    ChechAgentReadyCount = ChechAgentReadyCount + 1
    If ChechAgentReadyCount >= 60 Then '¨C0.5¬íÀË¬d¤@¦¸,¶W¹L30¬í
        tmrChkAgentReady.Enabled = False
        ShowError "PingAgent¹O®É¥¼¦^À³Ready!"
    End If
End Sub


Public Sub SaveEvent()
'Àx¦s¸ê°T

    Dim rsUpDown As ADODB.Recordset
    Dim i As Integer
    Dim itemx As ListItem
    Dim x As Long
    Dim eventtype As Long
    'Dim sentpkt As Integer
    Dim key As String
    Dim NodeIndex As Long
    
    '¶}©lÀx¦s¸ê®Æ
    On Error GoTo ErrHandler

    x = SendMessage(lvPingList.hWnd, WM_SETREDRAW, 0, 0)
    Set rsUpDown = New ADODB.Recordset
    With rsUpDown
        .CursorLocation = adUseClient
        .Open "UpDown", ConnStr, adOpenDynamic, adLockOptimistic
        For i = 1 To EventBufCount
            .AddNew
            !LogTime = aryEventBuf2(i)
            NodeIndex = aryEventBuf1(1, i) - 1
            !NodeName = aryNodeName(NodeIndex)
            eventtype = aryEventBuf1(2, i)
            !Event = eventtype
            If eventtype = -1 Then 'up--> down
                key = "#" & NodeIndex
                Set itemx = lvPingList.ListItems(key)
                itemx.SubItems(16) = Format(aryEventBuf2(i), "yyyy/mm/dd Hh:Nn:Ss")
            End If
            .Update
        Next
        .Close
        
    End With
    Set rsUpDown = Nothing
    TotalEventCount = TotalEventCount + EventBufCount
    If TotalEventCount > MAX_TOTAL_EVENT_COUNT Then
        tmrLog.Enabled = False
        ShowError "Event Log Exceed...!"
    End If
    x = SendMessage(lvPingList.hWnd, WM_SETREDRAW, 1, 0)
    Exit Sub
ErrHandler:
    Call StopLog
    MsgBox "Àx¦s¸ê®Æ®É²£¥Í¤U¦C¿ù»~:" & vbCrLf & Err.Description, vbExclamation, MsgTitle
End Sub
Private Function GetLogFileName() As String
    Dim LogDate As String
    Dim LogFile As String

    LogDate = Format(Date, "yyyy-mm-dd") '¤é´Á¤]¥i¯à¦b¥¼Ãö¾÷ªº¨t²Î¤¤·|§ïÅÜ
    LogFile = App.Path & "\log\" & LogDate & ".txt"
    GetLogFileName = LogFile
End Function


Public Sub RefreshPingStatus()
    
    Dim itemx As ListItem
    Dim i As Long

    Dim x As Long

    'Dim sentpkt As Integer
    Dim key As String
    Dim recvpkt As Long
    Dim lostpkt As Long
    Dim sentpkt As Long
    Dim alertcount As Long
    Dim CycleInterval As Long
    Dim tmpvalue As Long
    Dim failcount As Long
    Dim cyclecount As Long
    
    Dim SuccessCount As Integer
    Dim DownCount As Integer
    Dim WarnCount As Integer
    
    On Error GoTo ErrHandler
    With lvPingList
    x = SendMessage(.hWnd, WM_SETREDRAW, 0, 0)

    
    For i = 0 To MaxNodeIndex
        key = "#" & i
        Set itemx = .ListItems(key)
        
        
        recvpkt = aryReportData(2, i)
        lostpkt = aryReportData(3, i)
        sentpkt = recvpkt + lostpkt
        CycleInterval = aryReportData(6, i)
        If sentpkt > 0 Then
        
            'received
            itemx.SubItems(8) = Right(SPACE3 & recvpkt, 3)
            'avg rtt
            If recvpkt > 0 Then
                itemx.SubItems(11) = Right(SPACE3 & Format(Round(aryReportData(4, i) / recvpkt, 1), "###0.0"), 6)
            Else
                itemx.SubItems(11) = ""
            End If
            'lost
            itemx.SubItems(9) = Right(SPACE3 & lostpkt, 3)
            'sentpkt
            itemx.SubItems(7) = Right(SPACE3 & sentpkt, 3)
            
            'packet loss%
            If lostpkt = 0 Then
                itemx.SubItems(10) = "   0.0"
            ElseIf lostpkt = sentpkt Then
                itemx.SubItems(10) = " 100.0"
            Else
                itemx.SubItems(10) = Right(SPACE3 & Format(Round(100 * lostpkt / sentpkt, 1), "##0.0"), 6)
            End If
            'statistics cycle
            itemx.SubItems(12) = Right(SPACE3 & Format(Round(CycleInterval / 1000, 1), "###0.0"), 6)

        Else
            itemx.SubItems(8) = ""
            itemx.SubItems(11) = ""
            itemx.SubItems(9) = ""
            itemx.SubItems(7) = ""
            itemx.SubItems(10) = ""
            itemx.SubItems(12) = ""
        End If
        
        'fail count
        failcount = aryReportData(7, i)
        If failcount > 0 Then
            itemx.SubItems(13) = Right(SPACE5 & failcount, 6)
        Else
            itemx.SubItems(13) = ""
        End If
        
        'ping cycle count
        cyclecount = aryReportData(8, i)
        If cyclecount > 0 Then
            itemx.SubItems(14) = Right(SPACE5 & cyclecount, 6)
        End If
        
        'alert count
        alertcount = aryReportData(5, i)
        If alertcount > 0 Then
            itemx.SubItems(15) = Right(SPACE5 & alertcount, 6)
        Else
            itemx.SubItems(15) = ""
        End If
        
        Select Case aryReportData(1, i)
        Case RESULT_SUCCESS
        'Green
            SuccessCount = SuccessCount + 1
            If aryLastLedResult(i) <> RESULT_SUCCESS Then
                itemx.Text = " " '¸m©ó¤¤¶¡
                itemx.SmallIcon = "green"
                aryLastLedResult(i) = RESULT_SUCCESS
            End If
        Case RESULT_DOWN
        'Red
            DownCount = DownCount + 1
            If aryLastLedResult(i) <> RESULT_DOWN Then
                itemx.Text = SPACE2
                itemx.SmallIcon = "red"
                aryLastLedResult(i) = RESULT_DOWN
            End If
        Case RESULT_WARN
        'Yellow
            WarnCount = WarnCount + 1
            If aryLastLedResult(i) <> RESULT_WARN Then
                itemx.Text = "" '²¾¨ì³Ì«e­±,¦]¬°³Ì©_©Ç
                itemx.SmallIcon = "yellow"
                aryLastLedResult(i) = RESULT_WARN
            End If
        End Select
    Next
    
    x = SendMessage(.hWnd, WM_SETREDRAW, 1, 0)
    statusbar.Panels(2).Text = "Success: " & SuccessCount & " ­Ó¸`ÂI"
    statusbar.Panels(3).Text = "Warn: " & WarnCount & " ­Ó¸`ÂI"
    statusbar.Panels(4).Text = "Down: " & DownCount & " ­Ó¸`ÂI"
    statusbar.Panels(5).Text = "Check: " & (SuccessCount + DownCount + WarnCount) & "­Ó¸`ÂI"
    End With
    Exit Sub
ErrHandler:
    MsgBox "RefreshPingStatus error!" & vbCrLf & Err.Description, vbExclamation, MsgTitle

End Sub


Public Sub AgentLedBlink(AgentID As Long, BlinkSwitch As Long)
    Dim idx As Long
    aryAgentBlinkTick(AgentID) = GetTickCount
    Select Case BlinkSwitch
    Case MSG_AGENT_BLINK_A 'ºñ¦â
        idx = AgentID - 1
        picAgent(idx).Picture = picLed(1).Picture
    Case MSG_AGENT_BLINK_B '¶À¦â
        idx = AgentID - 1
        picAgent(idx).Picture = picLed(3).Picture
    'Case MSG_AGENT_BLINK_S
    
    End Select

End Sub
Public Sub CollectorLedBlink(BlinkSwitch As Long)
    glCollectorBlinkTick = GetTickCount
    
    Select Case BlinkSwitch
    Case MSG_PINGCOLLECTOR_BLINK_A 'ºñ¦â
        picPingCollector.Picture = picLed(1).Picture
    Case MSG_PINGCOLLECTOR_BLINK_B '¶À¦â
        picPingCollector.Picture = picLed(3).Picture
    'Case MSG_AGENT_BLINK_S
    End Select
End Sub


Private Sub tmrCheckCollectorAlive_Timer()
    'ping collector¬O§_alive
    Dim curtick As Long
    Dim CollectorTick As Long
    
    Const KEEP_ALIVE_TIME_OUT As Long = 5000
    If Not glCollectorReady Then
        tmrCheckCollectorAlive.Enabled = False
        Exit Sub
    End If
    
    curtick = GetTickCount
    CollectorTick = TickDiff(glCollectorBlinkTick, curtick)
    If CollectorTick > KEEP_ALIVE_TIME_OUT Then
        PingCollectorIsOff
        tmrCheckCollectorAlive.Enabled = False
    End If

End Sub
Private Sub tmrCheckAgentAlive_Timer()
    'ÀË¬dping agent ©M  ping collector¬O§_³£alive
    Dim curtick As Long
    Dim AgentID As Long
    Dim CloseAgentCount As Integer
    
    Const KEEP_ALIVE_TIME_OUT As Long = 5000
    curtick = GetTickCount
    '¥ýÀË¬dping agent
    CloseAgentCount = 0
    For AgentID = 1 To glAgentCount
        If Not aryAgentReady(AgentID) Then
            CloseAgentCount = CloseAgentCount + 1
        Else
            If TickDiff(aryAgentBlinkTick(AgentID), curtick) > KEEP_ALIVE_TIME_OUT Then
                AgentIsOff AgentID
                CloseAgentCount = CloseAgentCount + 1
            End If
        End If
    Next
    If CloseAgentCount >= glAgentCount Then
        tmrCheckAgentAlive.Enabled = False
    End If
End Sub
Public Sub AgentSayHello(AgentID As Long)
    Dim idx As Long
    idx = AgentID - 1
    picAgent(idx).Picture = picLed(3).Picture 'Say Hello®É§ï¦¨¶À¦â
    picAgent(idx).Refresh
    
    TellAgentToLoadINI AgentID '--> agent¦^³ø
End Sub

Public Sub AgentReady(AgentID As Long)
    Dim idx As Long
    aryAgentBlinkTick(AgentID) = GetTickCount
    idx = AgentID - 1
    picAgent(idx).Picture = picLed(1).Picture 'Ready®É§ï¦¨ºñ¦â
    picAgent(idx).Refresh
    
    aryAgentReady(AgentID) = True
End Sub
Public Sub PingCollectorReady()
    '§YPingCollector Load INI OK! --> °õ¦æCreatePingAgent
    
    picPingCollector.Picture = picLed(1).Picture 'ºñ¦â
    picPingCollector.Refresh
    glCollectorReady = True
    glCollectorBlinkTick = GetTickCount
    tmrCheckCollectorAlive.Interval = 2000
    tmrCheckCollectorAlive.Enabled = True
    ShowStatus "Load PingAgent.exe ..."
    Dim OK As Boolean
    OK = CreatePingAgent
    If OK Then
        Call EnableCheckAgentReady
    End If
End Sub
Public Sub AgentSayGoodbye(AgentID As Long)
    Dim idx As Long
    idx = AgentID - 1
    picAgent(idx).Picture = picLed(2).Picture 'red¦â
    picAgent(idx).Refresh
    aryAgentHwnd(AgentID) = 0
    aryAgentReady(AgentID) = False
End Sub
Public Sub AgentIsOff(AgentID As Long)
    '°»´ú¤£¨ì,¤£¬OAgent¥D°Ê»¡GoodBye
    Dim idx As Long
    idx = AgentID - 1
    picAgent(idx).Picture = picLed(2).Picture '¬õ¦â
    picAgent(idx).Refresh
    aryAgentReady(AgentID) = False
End Sub
Public Sub PingCollectorSayGoodbye()
    picPingCollector.Picture = picLed(2).Picture 'Log Agent Say Hello®É§ï¦¨ºñ¦â --> ÅÜred¦â
    picPingCollector.Refresh
    glCollectorReady = False
    glPingCollectorHwnd = 0
End Sub
Public Sub PingCollectorIsOff()
    picPingCollector.Picture = picLed(2).Picture 'Log Agent Say Hello®É§ï¦¨ºñ¦â --> ÅÜred¦â
    picPingCollector.Refresh
    glCollectorReady = False
End Sub
Public Sub InitAgentLedToGrayColor()
    Dim i As Integer
    For i = 0 To MaxAgent - 1
        picAgent(i).Picture = picLed(0).Picture
    Next
    'frLedPanel.Refresh
    Picture1.Refresh
End Sub


Private Sub CopyArray(FromArray As Variant, ToArray As Variant)
    'Copies source array to dest array
    'ToArray should be dynamic array
    Dim l As Long, lUBound As Long, lLBound As Long
    If (Not IsArray(FromArray)) Or (Not IsArray(ToArray)) Then Exit Sub
    
    lLBound = LBound(FromArray)
    lUBound = UBound(FromArray)
    ReDim ToArray(lLBound To lUBound)
    For l = lLBound To lUBound
        ToArray(l) = FromArray(l)
    Next

End Sub

Private Sub tmrGetReportData_Timer()
    If glCollectorReady Then
        TellCollectorToReportPingResult
    End If
End Sub

Private Sub tmrLog_Timer()
     '--> SaveEvent
    Dim curtick As Long
    Dim tdiff
    curtick = GetTickCount
    tdiff = TickDiff(LastLogTick, curtick) - glRefreshCycle
    'MsgBox "curtick,lastlogtick,tdiff = " & curtick & " ," & LastLogTick & " ," & tdiff
    If tdiff > glRefreshCycle Then
        '·íTellConnectorSendEvent¤§«á,collector¥²©w·|¥ß¨è¦^À³,¦Ó¥B¬O¦btmrLogªºInterval¤§¶¡
        '¦ý¬Otimer·|¦³µy·L»~®t,¬G¥²¶·¥[¤W¤@®e³\­È
        'MsgBox "curtick,lastlogtick,tdiff = " & curtick & " ," & LastLogTick & " ," & tdiff
        ShowError "log®É¶¡¹O®É " & tdiff & "ms"
    End If
    Call TellCollectorSendEvent
End Sub

Private Sub txtFind_GotFocus()
    txtFind.SelStart = 0
    txtFind.SelLength = Len(txtFind.Text)
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call cmdFindFirst_Click
    End If
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
'¥h±¼«öEnterªºbeepÁn
    Select Case KeyAscii
    Case vbKeyReturn
        KeyAscii = 0
    End Select
End Sub

Private Sub txtGoTo_GotFocus()
    txtGoTo.SelStart = 0
    txtGoTo.SelLength = Len(txtGoTo.Text)
    txtGoTo.SetFocus
End Sub

Private Sub txtGoTo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call cmdGoTo_Click
    End If
End Sub

Private Sub txtGoTo_KeyPress(KeyAscii As Integer)
'¥h±¼«öEnterªºbeepÁn
    Select Case KeyAscii
    Case vbKeyReturn
        KeyAscii = 0
    End Select
End Sub
Public Sub PingCollectorSayHello()
    'MsgBox glPingCollectorHwnd
    If glPingCollectorHwnd <> 0 Then
        picPingCollector.Picture = picLed(3).Picture '¶À¦â
        picPingCollector.Refresh
        ShowStatus "Collector Load List..."
        TellCollectorToLoadList NumOfPingNode
    End If
        
End Sub
Private Sub EnableCheckAgentReady()
    ChechAgentReadyCount = 0
    tmrChkAgentReady.Interval = 500 '¨C0.5¬íÀË¬d¤@¦¸
    tmrChkAgentReady.Enabled = True
End Sub
Public Sub ShowStatus(msg As String)
    statusbar.Panels(6).Text = msg
End Sub
Public Sub ShowError(msg As String)
    statusbar.Panels(6).Text = msg
End Sub
Public Sub StopLog()
    tmrLog.Enabled = False
End Sub
Private Sub StartLog()
    TotalEventCount = 0
    LastLogTick = GetTickCount
    tmrLog.Interval = glRefreshCycle
    tmrLog.Enabled = True
End Sub
