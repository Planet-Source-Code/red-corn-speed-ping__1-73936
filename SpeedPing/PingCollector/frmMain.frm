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
   Visible         =   0   'False
   Begin VB.Timer tmrBlink 
      Enabled         =   0   'False
      Interval        =   1270
      Left            =   10110
      Top             =   3960
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Test3"
      Height          =   330
      Left            =   11250
      TabIndex        =   8
      Top             =   3405
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.PictureBox picToolbar 
      Align           =   1  '¹ï»ôªí³æ¤W¤è
      Appearance      =   0  '¥­­±
      BorderStyle     =   0  '¨S¦³®Ø½u
      ForeColor       =   &H80000008&
      Height          =   1440
      Left            =   0
      ScaleHeight     =   1440
      ScaleWidth      =   15240
      TabIndex        =   6
      Top             =   0
      Width           =   15240
      Begin VB.PictureBox Picture1 
         Appearance      =   0  '¥­­±
         BorderStyle     =   0  '¨S¦³®Ø½u
         ForeColor       =   &H80000008&
         Height          =   990
         Left            =   15
         ScaleHeight     =   990
         ScaleWidth      =   17280
         TabIndex        =   17
         Top             =   60
         Width           =   17280
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   1
            Left            =   60
            Picture         =   "frmMain.frx":030A
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   137
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":0D0C
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   136
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":170E
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   135
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":2110
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   134
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":2B12
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   133
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":3514
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   132
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":3F16
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   131
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":4918
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   130
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":531A
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   129
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":5D1C
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   128
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":671E
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   127
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":7120
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   126
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":7B22
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   125
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":8524
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   124
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":8F26
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   123
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":9928
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   122
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":A32A
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   121
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":AD2C
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   120
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":B72E
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   119
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":C130
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   118
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":CB32
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   117
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":D534
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   116
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":DF36
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   115
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":E938
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   114
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":F33A
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   113
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":FD3C
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   112
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":1073E
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   111
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":11140
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   110
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":11B42
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   109
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":12544
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   108
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":12F46
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   107
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":13948
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   106
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":1434A
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   105
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":14D4C
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   104
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":1574E
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   103
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":16150
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   102
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":16B52
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   101
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":17554
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   100
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":17F56
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   99
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":18958
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   98
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":1935A
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   97
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":19D5C
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   96
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":1A75E
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   95
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":1B160
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   94
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":1BB62
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   93
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":1C564
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   92
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":1CF66
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   91
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":1D968
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   90
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":1E36A
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   89
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":1ED6C
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   88
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":1F76E
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   87
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":20170
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   86
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":20B72
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   85
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":21574
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   84
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":21F76
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   83
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":22978
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   82
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":2337A
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   81
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":23D7C
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   80
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":2477E
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   79
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":25180
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   78
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":25B82
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   77
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":26584
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   76
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":26F86
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   75
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":27988
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   74
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":2838A
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":28D8C
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":2978E
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":2A190
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":2AB92
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   69
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":2B594
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":2BF96
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":2C998
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":2D39A
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   65
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":2DD9C
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":2E79E
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":2F1A0
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   62
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":2FBA2
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":305A4
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":30FA6
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":319A8
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":323AA
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":32DAC
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":337AE
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":341B0
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":34BB2
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":355B4
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":35FB6
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":369B8
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":373BA
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":37DBC
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":387BE
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":391C0
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":39BC2
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":3A5C4
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":3AFC6
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":3B9C8
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":3C3CA
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":3CDCC
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":3D7CE
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":3E1D0
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":3EBD2
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":3F5D4
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":3FFD6
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":409D8
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":413DA
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":41DDC
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":427DE
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":431E0
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":43BE2
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":445E4
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":44FE6
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":459E8
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":463EA
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":46DEC
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":477EE
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":481F0
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":48BF2
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   270
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
            Picture         =   "frmMain.frx":495F4
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   705
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
            Picture         =   "frmMain.frx":49FF6
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   270
            Width           =   240
         End
         Begin VB.PictureBox picAgent 
            Appearance      =   0  '¥­­±
            AutoSize        =   -1  'True
            BorderStyle     =   0  '¨S¦³®Ø½u
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   119
            Left            =   16875
            Picture         =   "frmMain.frx":4A9F8
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   705
            Width           =   240
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   1
            Left            =   60
            TabIndex        =   257
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   256
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   2
            Left            =   345
            TabIndex        =   255
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   3
            Left            =   345
            TabIndex        =   254
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   4
            Left            =   630
            TabIndex        =   253
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   5
            Left            =   630
            TabIndex        =   252
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   6
            Left            =   915
            TabIndex        =   251
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   7
            Left            =   915
            TabIndex        =   250
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   8
            Left            =   1200
            TabIndex        =   249
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   9
            Left            =   1200
            TabIndex        =   248
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   10
            Left            =   1485
            TabIndex        =   247
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   11
            Left            =   1485
            TabIndex        =   246
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   12
            Left            =   1770
            TabIndex        =   245
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   13
            Left            =   1770
            TabIndex        =   244
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   14
            Left            =   2055
            TabIndex        =   243
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   15
            Left            =   2055
            TabIndex        =   242
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   16
            Left            =   2340
            TabIndex        =   241
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   17
            Left            =   2340
            TabIndex        =   240
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   18
            Left            =   2625
            TabIndex        =   239
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   19
            Left            =   2625
            TabIndex        =   238
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   20
            Left            =   2910
            TabIndex        =   237
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   21
            Left            =   2910
            TabIndex        =   236
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   22
            Left            =   3195
            TabIndex        =   235
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   23
            Left            =   3195
            TabIndex        =   234
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   24
            Left            =   3480
            TabIndex        =   233
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   25
            Left            =   3480
            TabIndex        =   232
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   26
            Left            =   3765
            TabIndex        =   231
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   27
            Left            =   3765
            TabIndex        =   230
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   28
            Left            =   4050
            TabIndex        =   229
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   29
            Left            =   4050
            TabIndex        =   228
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   30
            Left            =   4335
            TabIndex        =   227
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   31
            Left            =   4335
            TabIndex        =   226
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   32
            Left            =   4620
            TabIndex        =   225
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   33
            Left            =   4620
            TabIndex        =   224
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   34
            Left            =   4905
            TabIndex        =   223
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   35
            Left            =   4905
            TabIndex        =   222
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   36
            Left            =   5190
            TabIndex        =   221
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   37
            Left            =   5190
            TabIndex        =   220
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   38
            Left            =   5475
            TabIndex        =   219
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   39
            Left            =   5475
            TabIndex        =   218
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   40
            Left            =   5760
            TabIndex        =   217
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   41
            Left            =   5760
            TabIndex        =   216
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   42
            Left            =   6045
            TabIndex        =   215
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   43
            Left            =   6045
            TabIndex        =   214
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   44
            Left            =   6330
            TabIndex        =   213
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   45
            Left            =   6330
            TabIndex        =   212
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   46
            Left            =   6615
            TabIndex        =   211
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   47
            Left            =   6615
            TabIndex        =   210
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   48
            Left            =   6900
            TabIndex        =   209
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   49
            Left            =   6900
            TabIndex        =   208
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   50
            Left            =   7185
            TabIndex        =   207
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   51
            Left            =   7185
            TabIndex        =   206
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   52
            Left            =   7470
            TabIndex        =   205
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   53
            Left            =   7470
            TabIndex        =   204
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   54
            Left            =   7755
            TabIndex        =   203
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   55
            Left            =   7755
            TabIndex        =   202
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   56
            Left            =   8040
            TabIndex        =   201
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   57
            Left            =   8040
            TabIndex        =   200
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   58
            Left            =   8325
            TabIndex        =   199
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   59
            Left            =   8325
            TabIndex        =   198
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   60
            Left            =   8610
            TabIndex        =   197
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   61
            Left            =   8610
            TabIndex        =   196
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   62
            Left            =   8895
            TabIndex        =   195
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   63
            Left            =   8895
            TabIndex        =   194
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   64
            Left            =   9180
            TabIndex        =   193
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   65
            Left            =   9180
            TabIndex        =   192
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   66
            Left            =   9465
            TabIndex        =   191
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   67
            Left            =   9465
            TabIndex        =   190
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   68
            Left            =   9750
            TabIndex        =   189
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   69
            Left            =   9750
            TabIndex        =   188
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   70
            Left            =   10035
            TabIndex        =   187
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   71
            Left            =   10035
            TabIndex        =   186
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   72
            Left            =   10320
            TabIndex        =   185
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   73
            Left            =   10320
            TabIndex        =   184
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   74
            Left            =   10605
            TabIndex        =   183
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   75
            Left            =   10605
            TabIndex        =   182
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   76
            Left            =   10890
            TabIndex        =   181
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   77
            Left            =   10890
            TabIndex        =   180
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   78
            Left            =   11175
            TabIndex        =   179
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   79
            Left            =   11175
            TabIndex        =   178
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   80
            Left            =   11460
            TabIndex        =   177
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   81
            Left            =   11460
            TabIndex        =   176
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   82
            Left            =   11745
            TabIndex        =   175
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   83
            Left            =   11745
            TabIndex        =   174
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   84
            Left            =   12030
            TabIndex        =   173
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   85
            Left            =   12030
            TabIndex        =   172
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   86
            Left            =   12315
            TabIndex        =   171
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   87
            Left            =   12315
            TabIndex        =   170
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   88
            Left            =   12600
            TabIndex        =   169
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   89
            Left            =   12600
            TabIndex        =   168
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   90
            Left            =   12885
            TabIndex        =   167
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   91
            Left            =   12885
            TabIndex        =   166
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   92
            Left            =   13170
            TabIndex        =   165
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   93
            Left            =   13170
            TabIndex        =   164
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   94
            Left            =   13455
            TabIndex        =   163
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   95
            Left            =   13455
            TabIndex        =   162
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   96
            Left            =   13740
            TabIndex        =   161
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   97
            Left            =   13740
            TabIndex        =   160
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   98
            Left            =   14025
            TabIndex        =   159
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   99
            Left            =   14025
            TabIndex        =   158
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   100
            Left            =   14310
            TabIndex        =   157
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   101
            Left            =   14310
            TabIndex        =   156
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   102
            Left            =   14595
            TabIndex        =   155
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   103
            Left            =   14595
            TabIndex        =   154
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   104
            Left            =   14880
            TabIndex        =   153
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   105
            Left            =   14880
            TabIndex        =   152
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   106
            Left            =   15165
            TabIndex        =   151
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   107
            Left            =   15165
            TabIndex        =   150
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   108
            Left            =   15450
            TabIndex        =   149
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   109
            Left            =   15450
            TabIndex        =   148
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   110
            Left            =   15735
            TabIndex        =   147
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   111
            Left            =   15735
            TabIndex        =   146
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   112
            Left            =   16020
            TabIndex        =   145
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   113
            Left            =   16020
            TabIndex        =   144
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   114
            Left            =   16305
            TabIndex        =   143
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   115
            Left            =   16305
            TabIndex        =   142
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   116
            Left            =   16590
            TabIndex        =   141
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   117
            Left            =   16590
            TabIndex        =   140
            Top             =   510
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   118
            Left            =   16875
            TabIndex        =   139
            Top             =   75
            Width           =   285
         End
         Begin VB.Label lblAgent 
            Alignment       =   2  '¸m¤¤¹ï»ô
            Caption         =   "123"
            Height          =   195
            Index           =   119
            Left            =   16875
            TabIndex        =   138
            Top             =   510
            Width           =   285
         End
      End
      Begin VB.TextBox txtGoTo 
         Alignment       =   2  '¸m¤¤¹ï»ô
         Height          =   285
         Left            =   5685
         TabIndex        =   15
         Top             =   1110
         Width           =   585
      End
      Begin VB.CommandButton cmdGoTo 
         Caption         =   "Go To"
         Height          =   330
         Left            =   6330
         TabIndex        =   14
         Top             =   1080
         Width           =   1080
      End
      Begin VB.CommandButton cmdFindNext 
         Caption         =   "Find Next"
         Height          =   330
         Left            =   3675
         TabIndex        =   13
         Top             =   1080
         Width           =   1080
      End
      Begin VB.CommandButton cmdFindFirst 
         Caption         =   "Find First"
         Height          =   330
         Left            =   2595
         TabIndex        =   11
         Top             =   1080
         Width           =   1080
      End
      Begin VB.TextBox txtFind 
         Height          =   285
         Left            =   975
         TabIndex        =   10
         Top             =   1110
         Width           =   1545
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Test1"
         Height          =   330
         Left            =   15345
         TabIndex        =   9
         Top             =   75
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.CommandButton cmdExcel 
         Caption         =   "Excel"
         Height          =   330
         Left            =   8745
         TabIndex        =   7
         Top             =   1080
         Width           =   1080
      End
      Begin VB.Label Label2 
         Alignment       =   2  '¸m¤¤¹ï»ô
         Caption         =   "Go to #:"
         Height          =   195
         Left            =   4980
         TabIndex        =   16
         Top             =   1155
         Width           =   720
      End
      Begin VB.Label Label1 
         Alignment       =   2  '¸m¤¤¹ï»ô
         Caption         =   "Find Node:"
         Height          =   195
         Left            =   75
         TabIndex        =   12
         Top             =   1155
         Width           =   900
      End
   End
   Begin VB.Timer tmrRefreshList 
      Enabled         =   0   'False
      Left            =   9015
      Top             =   3945
   End
   Begin VB.PictureBox picLed 
      Height          =   315
      Index           =   0
      Left            =   7890
      Picture         =   "frmMain.frx":4B3FA
      ScaleHeight     =   255
      ScaleWidth      =   240
      TabIndex        =   5
      Top             =   4605
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picLed 
      Height          =   315
      Index           =   1
      Left            =   8295
      Picture         =   "frmMain.frx":4BDFC
      ScaleHeight     =   255
      ScaleWidth      =   240
      TabIndex        =   4
      Top             =   4605
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picLed 
      Height          =   315
      Index           =   2
      Left            =   8655
      Picture         =   "frmMain.frx":4C7FE
      ScaleHeight     =   255
      ScaleWidth      =   240
      TabIndex        =   3
      Top             =   4605
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picLed 
      Height          =   315
      Index           =   3
      Left            =   9015
      Picture         =   "frmMain.frx":4D200
      ScaleHeight     =   255
      ScaleWidth      =   240
      TabIndex        =   2
      Top             =   4605
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Timer tmrPingQueue 
      Enabled         =   0   'False
      Left            =   8430
      Top             =   3960
   End
   Begin MSComctlLib.ListView lvPingList 
      Height          =   2610
      Left            =   225
      TabIndex        =   0
      Top             =   2550
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
      NumItems        =   15
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
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "IP Address"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Min RTT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Avg RTT"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   6
         Text            =   "Max RTT"
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
         Text            =   "Ping Cycle(sec)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   12
         Text            =   "Continued Fail"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   13
         Text            =   "UpDown Alert"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Status Code"
         Object.Width           =   2540
      EndProperty
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
            Picture         =   "frmMain.frx":4DC02
            Key             =   "green"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4E614
            Key             =   "red"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4F026
            Key             =   "yellow"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4FA38
            Key             =   "gray"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar statusbar 
      Align           =   2  '¹ï»ôªí³æ¤U¤è
      Height          =   330
      Left            =   0
      TabIndex        =   1
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
            Picture         =   "frmMain.frx":5044A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5051C
            Key             =   ""
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
Private PingQueueIndex As Integer

Private CheckIsRunning As Boolean
Private LastCheckEnd As Long

Private Const WM_SETREDRAW = &HB
Private Const LVM_FIRST As Long = &H1000
Private Const LVM_SETCOLUMNWIDTH As Long = (LVM_FIRST + 30)
Private Const LVSCW_AUTOSIZE As Long = -1
Private Const LVSCW_AUTOSIZE_USEHEADER As Long = -2

Private ListViewSortOrder() As Integer

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long

Private Const WM_QUIT = &H12
Private Const WM_CLOSE = &H10
Private Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hwnd As Long, _
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
                            (ByVal hwnd&, RCT As RECT)
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Private SPACE9 As String ''¸ü¤Jfailcount ,successcount¥Î(listview±Æ§Ç)
Private SPACE5 As String
Private SPACE3 As String

Private LogCounter As Integer

Private ChechAgentReadyCount As Long
Private CollectorBlinkSwitch As Integer
Private Sub cmdExcel_Click()
    Dim excl As Excel.Application
    Dim bk As Workbook
    Dim sht As Worksheet
    Dim i As Integer, j As Integer
    Dim txtcontent As String
    Dim tmpvalue As String
    Dim tmpline As String
    Dim status As String
    
    On Error GoTo ErrHandler
    Me.MousePointer = vbHourglass
    txtcontent = "Status"
    For i = 2 To lvPingList.ColumnHeaders.Count
        txtcontent = txtcontent & vbTab & lvPingList.ColumnHeaders(i)
    Next i
    
    For i = 1 To lvPingList.ListItems.Count
        status = lvPingList.ListItems(i).Text
        Select Case status
        Case ""
          tmpline = "Y"
        Case " "
          tmpline = "G"
        Case "  "
          tmpline = "R"
        End Select
        
        For j = 2 To lvPingList.ColumnHeaders.Count
            tmpvalue = Trim(lvPingList.ListItems(i).SubItems(j - 1))
            Select Case j
            Case 3, 5 'node name
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
    findstr = Trim(txtFind)
    If findstr = "" Then
        Exit Sub
    End If
    For i = 1 To lvPingList.ListItems.Count
        'Set itmX = lvPingList.ListItems(i).SubItems(2)
        If InStr(1, lvPingList.ListItems(i).SubItems(2), findstr, vbTextCompare) > 0 Then
            Set itmX = lvPingList.ListItems(i)
            itmX.EnsureVisible
            itmX.Selected = True
            lvPingList.SetFocus
            Exit Sub
        End If
    Next
    '¥Î¦¹¤èªk(lvwPartial)¥u¯à·j´MlvwText
    'Set itmX = lvPingList.FindItem(findstr, lvwSubItem, , lvwPartial)
    
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
    findstr = Trim(txtFind)
    If findstr = "" Then
        Exit Sub
    End If
    findstart = lvPingList.SelectedItem.Index + 1
    If findstart > lvPingList.ListItems.Count Then
        findstart = 1
    End If
    For i = findstart To lvPingList.ListItems.Count
        'Set itmX = lvPingList.ListItems(i).SubItems(2)
        If InStr(1, lvPingList.ListItems(i).SubItems(2), findstr, vbTextCompare) > 0 Then
            Set itmX = lvPingList.ListItems(i)
            itmX.EnsureVisible
            itmX.Selected = True
            lvPingList.SetFocus
            Exit Sub
        End If
    Next
    '¥Î¦¹¤èªk(lvwPartial)¥u¯à·j´MlvwText
    'Set itmX = lvPingList.FindItem(findstr, lvwSubItem, , lvwPartial)
    MsgBox "§ä¤£¨ì¦¹¸`ÂI!", vbExclamation, MsgTitle
    txtFind.SelStart = 0
    txtFind.SelLength = Len(txtFind.Text)
    txtFind.SetFocus
End Sub

Private Sub cmdGoTo_Click()
    Dim itmX As ListItem
    Dim num As Integer
    Dim i As Integer
    
    If IsNumeric(txtGoTo.Text) Then
        num = CInt(txtGoTo.Text)
        If num > lvPingList.ListItems.Count Then
            num = lvPingList.ListItems.Count
        End If
        If num < 1 Then
            num = 1
        End If

        'Set itmX = lvPingList.ListItems(num)
        '§ï¥Î³oºØ¤èªk,¦]¬°itemªº¦ì¸m¥i¯à¦]¬°Äæ¦ìªº±Æ§Ç¦Ó¤£¦P©ósn
        For i = 1 To lvPingList.ListItems.Count
            If lvPingList.ListItems(i).Tag = num Then
                Set itmX = lvPingList.ListItems(i)
                itmX.EnsureVisible
                itmX.Selected = True
                lvPingList.SetFocus
                Exit Sub
            End If
        Next
    Else
        txtGoTo.SelStart = 0
        txtGoTo.SelLength = Len(txtGoTo.Text)
        txtGoTo.SetFocus
    End If
End Sub

Private Sub InitListViewSort()
    Dim col As ColumnHeader
    Set col = lvPingList.ColumnHeaders(2)     'sort on first column
    lvPingList.ColumnHeaderIcons = imlSortIcon
    ListViewSortOrder(2) = lvwDescending      'will get flipped to ascending
    lvPingList_ColumnClick col                'click the column heading
    lvPingList.ListItems(1).EnsureVisible     'make sure first one is visible
    lvPingList.ListItems(1).Selected = True   'and selected
End Sub

'Private Sub cmdLoadList_Click()
Public Sub DoLoadPingList(num As Long)
    Dim loadok As Boolean

    loadok = LoadListFromDBIntoArray
    If Not loadok Then
        Exit Sub
    End If
    If NumOfPingNode <> num Then
        Call ReportLoadListERR
    End If
    If NumOfPingNode > 0 Then

        ReDim aryPingStatData(1 To 8, MaxNodeIndex)
        ReDim aryContinuedFail(MaxNodeIndex)
        ReDim aryReportData(1 To 8, MaxNodeIndex)
        ReDim aryLastLedResult(MaxNodeIndex)
        ReDim aryAgentPingResultData(1 To 6, MaxNodeIndex)
        
        ReDim aryLastUpDown(MaxNodeIndex)
        LoadPingListIntoListView
        
        statusbar.Panels(2).Text = ""
        statusbar.Panels(3).Text = ""
        statusbar.Panels(4).Text = ""
        statusbar.Panels(5).Text = ""
        ReDim ListViewSortOrder(1 To lvPingList.ColumnHeaders.Count)
        Call InitListViewSort
        Call ReportLoadListOK
    End If
End Sub

Private Function LoadListFromDBIntoArray() As Boolean
    Dim rsPingList As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo ErrHandler
    LoadListFromDBIntoArray = False
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
        ReDim arySN(MaxNodeIndex)
        .MoveFirst
        i = 0
        While Not .EOF
            arySN(i) = !SN
            aryNodeName(i) = !NodeName
            aryIPAddress(i) = !IP
            aryInetAddr(i) = inet_addr(aryIPAddress(i))
            i = i + 1
            .MoveNext
        Wend
        .Close
        Set rsPingList = Nothing
    End With
    LoadListFromDBIntoArray = True
    Exit Function
ErrHandler:
    MsgBox "¥ÑDB¸ü¤Jping list®Éµo¥Í¿ù»~:" & vbCrLf & Err.Description, vbExclamation, MsgTitle
End Function


Public Sub StartPing()
    If NumOfPingNode = 0 Then
        Exit Sub
    End If
    
    Dim i As Long

    ReDim aryNodeLastStatisticsCycleTick(MaxNodeIndex)

    Call ResetEventBuf
    
    statusbar.Panels(2).Text = ""
    statusbar.Panels(3).Text = ""
    statusbar.Panels(4).Text = ""
    statusbar.Panels(5).Text = ""
    
    
    '³°Äò±Ò°Ê¨C­Óagent,¦Ó¤£¤@¦¸±Ò°Ê
    
    PingQueueIndex = 0
    tmrPingQueue.Interval = 5
    tmrPingQueue.Enabled = True
    PingIsRunning = True
    If glDebugMode Then
        tmrRefreshList.Interval = glRefreshCycle
        tmrRefreshList.Enabled = True
    Else
        tmrRefreshList.Enabled = False
    End If
    
End Sub


Public Sub StopPing()

    'tmrSummary.Enabled = False
    PingIsRunning = False
    tmrRefreshList.Enabled = False
    tmrPingQueue.Enabled = False
    
    'Call SaveLog
End Sub


Private Sub Form_Load()
    On Error GoTo ErrHandler
    
    Me.Caption = MsgTitle

    glMyHwnd = Me.hwnd
    glMyPingMsg = RegisterWindowMessage(MSG_MYPING_POSTMSG)
    If glMyPingMsg = 0 Then
        MsgBox "Windows µLªkµù¥UMSG_MYPING_POSTMSG! µ{¦¡³q°T±N¨ü¨ì¼vÅT!", vbExclamation, MsgTitle
    End If
    lvPingList.Left = 45
    lvPingList.Top = picToolbar.Height + 10
    
    Dim i As Integer
    For i = 0 To MaxAgent - 1
        lblAgent(i).Caption = i + 1
    Next
    Call InitAgentLed
    
    SPACE9 = Space(9)
    SPACE5 = Space(5)
    SPACE3 = Space(3) '¥H¤W¬Ò¬O¬°¤Flistview±Æ§Ç¥Î,¤£¬O¬°¤FÅã¥Ü,¦]¬°Åã¥Ü¤W°£¤F©T©w¤j¤p¦rÅé¥~,·|¦]¦r«¬ªº½t¬G,¦r¤¸·|¦³¤j¤p¤§¤À
    
    
    'log path
    Dim fs As New FileSystemObject
    Dim txtfile As TextStream
    
    Dim PingLogDBFile As String
    PingLogDBFile = App.Path & "\..\pinglog.mdb"
    PingLogDBFile = fs.GetAbsolutePathName(PingLogDBFile)
    
    If Not fs.FileExists(PingLogDBFile) Then
        MsgBox PingLogDBFile & " ¸ê®Æ®wÀÉ®×¤£¦s¦b!", vbCritical, MsgTitle
        End
    End If
   
    ConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & PingLogDBFile & ";Persist Security Info=False"
    
    Call Hook '­n¦³Hook, copy memory¤~¯à§@¥Î
    
    'Åª¨ú©R¥O¦C°Ñ¼Æ
    Dim cmdline As String
    Dim aryPara() As String
    'Dim waittmr As ClassWaitableTimer
    cmdline = Trim(Command)
    If cmdline <> "" Then
        aryPara = Split(cmdline, Space(1))
        If UBound(aryPara) <> 1 Then
            MsgBox "¸ü¤JPing Collectorµ{¦¡®Éµo¥Í°Ñ¼Æ¼Æ¥Ø¿ù»~!", vbCritical, MsgTitle
            Unload Me
        End If
        glPingManagerHwnd = aryPara(0)
        glDebugMode = aryPara(1)
        If glDebugMode Then
            Me.Visible = True
        Else
            Me.Visible = False
            App.TaskVisible = False
        End If
'        Set waittmr = New ClassWaitableTimer
'        waittmr.Wait 200
        Call SayHello
'        Set waittmr = Nothing
        'MsgBox glDebugMode
        
    Else
        '§ï¥Ñ¸ê®Æ®w¸ü¤J
        glDebugMode = True
        Me.Visible = True
        glPingManagerHwnd = 0 '¤£¥Î¦^³ø
    End If
    Exit Sub
ErrHandler:
    MsgBox Err.Description, vbCritical, MsgTitle
    End
End Sub
Public Sub StartBlink()
    tmrBlink.Interval = 1070
    tmrBlink.Enabled = True
End Sub

Private Sub Form_Resize()
    Dim clientrect As RECT
    On Error GoTo ErrHandler
    If Me.WindowState = vbMinimized Then Exit Sub
    GetClientRect Me.hwnd, clientrect
    Picture1.Width = (clientrect.Right - clientrect.Left) * Screen.TwipsPerPixelX - 120
    'Bevel1.Width = Picture1.Width - 120
'    If frLedPanel.Width > picLeds.Width Then
'        picLeds.Left = (frLedPanel.Width - picLeds.Width) / 2
'    Else
'        picLeds.Left = 45
'    End If
    'ª`·N:¤£­n¥Îme.Height,¤@¥¹¤Á´«classic style / xp style¤£¦Pªºcontrol®É,¶ZÂ÷·|¦³»~®t
    lvPingList.Height = (clientrect.Bottom - clientrect.Top) * Screen.TwipsPerPixelY - picToolbar.Height - statusbar.Height - 25
    lvPingList.Width = (clientrect.Right - clientrect.Left) * Screen.TwipsPerPixelX - 90
   
    Call AdjustColWidth
    Exit Sub
ErrHandler:

End Sub
Private Sub LoadPingListIntoListView(Optional IsInit As Boolean = False)
    Dim X As Long
    Dim i As Long
    Dim j As Long
    Dim itemx As ListItem
    Dim SN As String
    Dim sbuffer As String
    Dim maxnumlen As Integer
    On Error GoTo ErrHandler
    X = SendMessage(lvPingList.hwnd, WM_SETREDRAW, 0, 0)
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
                
                
            Next
        End If

'        ColorGrid lvPingList
    End With
    X = SendMessage(lvPingList.hwnd, WM_SETREDRAW, 1, 0)
    Call AdjustColWidth '­n©ñ¦bWM_SETREDRAW¤§«á,§_«h¹Ï¥Ü·|¦³´Ý¼v
    statusbar.Panels(1).Text = "Á`¦@ " & NumOfPingNode & " ¸`ÂI"
    Exit Sub

ErrHandler:
        'Screen.MousePointer = vbDefault
        MsgBox "¸ü¤JPing List®Éµo¥Í¿ù»~!" & vbCrLf & Err.Description, vbCritical, MsgTitle
        Exit Sub
End Sub

Private Sub AdjustColWidth()
    'Size each column based on the maximum of
  'EITHER the column header text width, or,
  'if the items below it are wider, the
  'widest list item in the column.
  '
  'The last column is always resized to occupy
  'the remaining width in the control.
    Dim startcol As Long
    Dim col2adjust As Long
    If lvPingList.View = lvwReport Then
        
    '¦Û°Ê½Õ¾ãÄæ¼e
        'lvPingList.ColumnHeaders(1).Width = 300
        If lvPingList.ColumnHeaders(1).Text = "" Then
            startcol = 1
        Else
            startcol = 0
        End If
        For col2adjust = startcol To lvPingList.ColumnHeaders.Count - 1
            'If col2adjust <> 1 Then 'ÁôÂÃÄæ¦ì
                Call SendMessage(lvPingList.hwnd, _
                     LVM_SETCOLUMNWIDTH, _
                     col2adjust, _
                     ByVal LVSCW_AUTOSIZE_USEHEADER)
            'End If
        Next
        
    End If
End Sub




Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Call SayGoodbye
    tmrPingQueue.Enabled = False
    tmrBlink.Enabled = False
    tmrRefreshList.Enabled = False
    tmrBlink.Enabled = False
    UnHook
    Dim frm    As Form
    For Each frm In Forms
        Unload frm
    Next
End Sub


Private Sub lvPingList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
' °õ¦æ±Æ§Ç
' 1 is the first column
    Dim i As Integer
    With lvPingList
        If lvPingList.ListItems.Count = 0 Then Exit Sub
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
        ShowListViewSortIcon lvPingList
        
'        If Not .SelectedItem Is Nothing Then
'            .SelectedItem.EnsureVisible
'        End If
        .ListItems(1).EnsureVisible     'make sure first one is visible
        .ListItems(1).Selected = True   'and selected
        
    End With
    
End Sub

Public Sub SetRefreshRange(AgentID As Long, NodeIndexHigh As Long)
    aryAgentRefreshRange(AgentID, 2) = NodeIndexHigh '¥u°O¿ý·sªºÅã¥ÜºI¤î¦ì¸m§Y¥i
End Sub

Private Sub tmrBlink_Timer()
    If PingIsRunning Then
        CollectorBlinkSwitch = IIf(CollectorBlinkSwitch = 0, 1, 0)
        CollectorLedBlink CollectorBlinkSwitch
        Exit Sub
    End If
    
    If CollectorBlinkSwitch = 2 Then
        CollectorLedBlink CollectorBlinkSwitch
    Else
        If CollectorBlinkSwitch = 0 Then '¦³¥i¯à¤w¸g¬Oºñ¦â
            CollectorBlinkSwitch = 2
            CollectorLedBlink CollectorBlinkSwitch
        Else
            CollectorLedBlink 0 '¥ý°{¦¨ºñ¦â
            CollectorBlinkSwitch = 2
        End If
    End If
End Sub

Private Function ChkAgentReady() As Boolean
    Dim i As Integer
    For i = 1 To glAgentCount
        If Not aryAgentReady(i) Then
            ChkAgentReady = False
            Exit Function
        End If
    Next
    ChkAgentReady = True
End Function

Private Sub SaveLog()
    '***°Olog
    '***¦pªGLogBuf¬°ªÅ«h¤£¥Î°Olog
    If EventBufCount = 0 Then
        Exit Sub
    End If
    'Call SendEventDataToLogAgent
    Call ResetEventBuf
    
    Exit Sub
    '**************************************************
    If LogBuf = "" Then
        Exit Sub
    End If
    Dim fs As FileSystemObject
    Dim F As TextStream
    '
    On Error GoTo ErrHandler
    Set fs = New FileSystemObject
    Set F = fs.OpenTextFile(GetLogFileName, ForAppending, True)
    F.Write LogBuf
    F.Close
    LogBuf = "" '¼g§¹¤§«á¥ß§Y²MªÅbuffer
    Set F = Nothing
    Set fs = Nothing
    
    Exit Sub
ErrHandler:
    MsgBox "°OLogÀÉ®Éµo¥Í¤F¥H¤Uªº¿ù»~!" & vbCrLf & Err.Description, vbExclamation, MsgTitle
End Sub
Private Function GetLogFileName() As String
    Dim LogDate As String
    Dim LogFile As String

    LogDate = Format(Date, "yyyy-mm-dd") '¤é´Á¤]¥i¯à¦b¥¼Ãö¾÷ªº¨t²Î¤¤·|§ïÅÜ
    LogFile = App.Path & "\log\" & LogDate & ".txt"
    GetLogFileName = LogFile
End Function
Private Sub tmrRefreshList_Timer()
    Dim AgentID As Long
    Dim lb As Long
    Dim ub As Long
    Dim X As Long
    On Error GoTo ErrHandler
    If Not glDebugMode Then Exit Sub
    X = SendMessage(lvPingList.hwnd, WM_SETREDRAW, 0, 0)
    For AgentID = 1 To glAgentCount
        '¥ýÅª¶i¨ìÅÜ¼Æ,¥H¨¾¦b§ó·slist®É,¤º®e­È³Q§ó§ï
        lb = aryAgentRefreshRange(AgentID, 1)
        ub = aryAgentRefreshRange(AgentID, 2)
        If ub >= 0 Then 'refresh end
            If lb > ub Then
                ub = aryAgentUB(AgentID) '¤w¸g¶}©l·sªº¤@½ü¤F,¥ýÅã¥Ü¨ìµ²§À,¨ä¥¦ªº¤U¤@¦¸timer®É¦AÅã¥Ü
            End If
            
            RefreshPingStatus lb, ub
            
            '­pºâ·sªºlb
            lb = ub + 1
            If lb > aryAgentUB(AgentID) Then
                aryAgentRefreshRange(AgentID, 1) = aryAgentLB(AgentID) '°O¿ý·sªºrefresh start
                aryAgentRefreshRange(AgentID, 2) = -1 '°O¿ý·sªºrefresh end
            Else
                aryAgentRefreshRange(AgentID, 1) = lb '°O¿ý·sªºrefresh start
                aryAgentRefreshRange(AgentID, 2) = -1 '°O¿ý·sªºrefresh end
            End If
        End If
    Next
    
    X = SendMessage(lvPingList.hwnd, WM_SETREDRAW, 1, 0)
    
    Dim SuccessCount As Integer
    Dim DownCount As Integer
    Dim WarnCount As Integer
    Dim test101 As Integer
    Dim i As Long
    Dim status As Long
    
    test101 = 0
    SuccessCount = 0
    DownCount = 0
    WarnCount = 0
    For i = 0 To MaxNodeIndex
        If aryAgentPingResultData(6, i) = PING_NOT_YET Then
            test101 = test101 + 1
        End If
        Select Case aryPingStatData(1, i)
        Case RESULT_DOWN
            DownCount = DownCount + 1
        Case RESULT_WARN
            WarnCount = WarnCount + 1
        Case RESULT_SUCCESS
            SuccessCount = SuccessCount + 1
        End Select
        
    Next
    statusbar.Panels(2).Text = "Success: " & SuccessCount & " ­Ó¸`ÂI"
    statusbar.Panels(3).Text = "Warn: " & WarnCount & " ­Ó¸`ÂI"
    statusbar.Panels(4).Text = "Down: " & DownCount & " ­Ó¸`ÂI"
    statusbar.Panels(5).Text = "Check: " & (SuccessCount + DownCount + WarnCount) & "­Ó¸`ÂI"
    If test101 > 0 Then
      statusbar.Panels(6).Text = "(Status Code 101): " & test101 & " ­Ó¸`ÂI"
    End If
    
    
    Exit Sub
ErrHandler:
    MsgBox "agentid=" & AgentID & ", lb=" & lb & ", ub=" & ub & ", agentlb=" & aryAgentLB(AgentID)
End Sub
Public Sub UpdatePingStatus(AgentID As Long, CheckStart As Long, CheckEnd As Long, cycletick As Long)
    Dim EventTime As String
    Dim FailedTime As String
    Dim t As Date
    Dim i As Long
    Dim j As Integer

    Dim SuccessCount As Integer
    Dim FailCount As Integer
    Dim status As Long
    Dim tick1 As Long
    
    'If (AgentID <> 24) Then Exit Sub
    
    If CheckIsRunning Then
        Exit Sub
    End If
    CheckIsRunning = True
'    MsgBox "CheckStart=" & CheckStart
'    MsgBox "CheckEnd=" & CheckEnd
    t = Now
    EventTime = Format(t, "Hh:Nn:Ss")
    FailedTime = Format(t, "mm/dd Hh:Nn:Ss")
    tick1 = GetTickCount
    For i = CheckStart To CheckEnd
        'ping cycle + 1
        aryPingStatData(8, i) = aryPingStatData(8, i) + 1
        
        aryPingStatData(6, i) = cycletick
         
'        'acc received packet
'        aryPingStatData(2, i) = aryPingStatData(2, i) + aryAgentPingResultData(4, i)
'        'acc lost pkt
'        aryPingStatData(3, i) = aryPingStatData(3, i) + aryAgentPingResultData(5, i)
'        'acc rtt
'        aryPingStatData(4, i) = aryPingStatData(4, i) + aryAgentPingResultData(2, i)
        'acc received packet
        aryPingStatData(2, i) = aryAgentPingResultData(4, i)
        'acc lost pkt
        aryPingStatData(3, i) = aryAgentPingResultData(5, i)
        'acc rtt
        aryPingStatData(4, i) = aryAgentPingResultData(2, i)
        '§PÂ_´ú¸Õ¦¨¥\»P§_
        If aryAgentPingResultData(6, i) <> MY_PING_OK Then
            '***´ú¸Õ¥¢±Ñ
            aryPingStatData(7, i) = aryPingStatData(7, i) + 1
            '¥þ­±¨ú®ø¦¹´ú¸Õ,¤@­Óµ{¦¡¤@ª½run¬O¤£¥i¯àªº,Àô¹Ò·|§ïÅÜ(iplist)If aryFailCount(i) < MAX_LONG_VALUE Then
            'ping¥¢±Ñ²Ö­p¦¸¼Æ(°O¦í,­n¹F¨ìcycle¤~¯àÂ_©w¬O§_¬°down)
            aryContinuedFail(i) = aryContinuedFail(i) + 1 'continued fail
            If aryContinuedFail(i) < glContinuedFailAsDown Then
                '¶È¹F¨ìÄµ§i,¥¼¹F§iÄµ
                aryPingStatData(1, i) = RESULT_WARN
            ElseIf aryContinuedFail(i) = glContinuedFailAsDown Then
                '½T©w¬°Down¤F,¦pªG¤W¤@¦¸¤£¬OUnknownªº¸Ü­nµo§iÄµ,¤]´N¬O¤W¤@¦¸ÁÙ¦n¦nªº
                If (aryLastUpDown(i) = RESULT_SUCCESS) Then
                    '§iÄµ¦¸¼Æ+1
                    aryPingStatData(5, i) = aryPingStatData(5, i) + 1
                    
                    '***°Olog***
                    EventBufCount = EventBufCount + 1
                    If EventBufCount >= MAX_EVENT_BUF Then
                        statusbar.Panels(1).Text = "Event buffer ¶W¹L³Ì¤j­È" & MAX_EVENT_BUF
                        Call ResetEventBuf
                    End If
                    
                    aryEventBuf1(1, EventBufCount) = arySN(i)
                    aryEventBuf1(2, EventBufCount) = -1 'Up-->Down
                    aryEventBuf2(EventBufCount) = t
                    
                    LogBuf = LogBuf & EventTime & SPACE3 & arySN(i) & SPACE3 & aryNodeName(i) & "   Down" & vbCrLf
                End If
                '±µ¤U¨Ó¤W¤@¦¸´N·|¬ODownªº
                aryLastUpDown(i) = RESULT_DOWN
                aryPingStatData(1, i) = RESULT_DOWN
            'Else 'aryContinuedFail(i) > glContinuedFailAsDown
            '   «O«ù¬°Down,¤£¥Î°Olog
                
            End If
        Else
            '***´ú¸Õ¦¨¥\
            
            aryContinuedFail(i) = 0 '¥¢±Ñ¦¸¼Æ­«·sÂk¹s 'continued fail
            
            If aryLastUpDown(i) = RESULT_DOWN Then
                '***°Olog,¦ý¤£¬O§iÄµ®@! ¦]¬°¬O¥ÑDown --> Up ***
                EventBufCount = EventBufCount + 1
                If EventBufCount >= MAX_EVENT_BUF Then
                    statusbar.Panels(1).Text = "Event buffer ¶W¹L³Ì¤j­È" & MAX_EVENT_BUF
                    Call ResetEventBuf
                End If
                
                aryEventBuf1(1, EventBufCount) = arySN(i)
                aryEventBuf1(2, EventBufCount) = 1 'Down-->Up
                aryEventBuf2(EventBufCount) = t
                
                LogBuf = LogBuf & EventTime & SPACE3 & arySN(i) & SPACE3 & aryNodeName(i) & "   Up" & vbCrLf
                
            End If
            aryLastUpDown(i) = RESULT_SUCCESS '°O¿ý´¿¸gUP¹L
            aryPingStatData(1, i) = RESULT_SUCCESS
        End If
        '²Î­pcycle(¿ù»~:¦]¬°continued fail¨Ã¤£¬O¨C¶¡¹j¤@©wªºcycle,¦Ó¬O¨C¤@¦¸¬Ò¥i¯à¹F¨ìcontinued as fail

'        If aryPingStatData(8, i) = glStatisticsCycle Then
'
'            '±N²Î­p­È½Æ»s¨ìreport data
'            aryReportData(1, i) = aryPingStatData(1, i) 'result
'            aryReportData(2, i) = aryPingStatData(2, i) 'received
'            aryReportData(3, i) = aryPingStatData(3, i) 'lost
'            aryReportData(4, i) = aryPingStatData(4, i) 'acc rtt
'
'            aryReportData(5, i) = TickDiff(aryNodeLastStatisticsCycleTick(i), tick1) 'cyc interval
'            aryNodeLastStatisticsCycleTick(i) = tick1
'            aryReportData(6, i) = aryPingStatData(5, i) 'alert count
'            aryReportData(7, i) = aryReportData(7, i) + 1 'stat. cycle count
'            If aryReportData(1, i) = RESULT_DOWN Then 'stat. fail count
'                aryReportData(8, i) = aryReportData(8, i) + 1
'            End If
'
'            '±N²Î­p­ÈÂk0
'            aryPingStatData(2, i) = 0 'acc received
'            aryPingStatData(3, i) = 0 'acc lost
'            aryPingStatData(4, i) = 0 'acc rtt
'            'aryPingStatData(5, i) = 0 'alert count Up-->Down,¦¸¼Æ¤£Âk¹s
'            aryPingStatData(8, i) = 0 'Ping Cycle count
'        End If
    Next
    SetRefreshRange AgentID, CheckEnd
    
    CheckIsRunning = False
    
End Sub

Private Sub RefreshPingStatus(CheckStart As Long, CheckEnd As Long)
    
    Dim itemx As ListItem
    Dim i As Long
    Dim j As Integer
    Dim maxrtt As Long, avgrtt As Single, minrtt As Long, sumrtt As Single
    Dim X As Long
    Dim recvpkt As Integer
    Dim lostpkt As Integer
    'Dim sentpkt As Integer
    Dim key As String
    Dim SuccessCount As Integer
    Dim FailCount As Integer
    Dim status As Long
    
    On Error GoTo ErrHandler
    With lvPingList
    'x = SendMessage(.hwnd, WM_SETREDRAW, 0, 0)

    
    For i = CheckStart To CheckEnd
        key = "#" & i
        'Set itemx = .ListItems(i + 1)
        Set itemx = .ListItems(key)
        recvpkt = aryAgentPingResultData(4, i)
        lostpkt = aryAgentPingResultData(5, i)
        'sentpkt = recvpkt + lostpkt
        
        'min rtt
        If aryAgentPingResultData(1, i) >= 0 Then
            itemx.SubItems(4) = Right(SPACE5 & aryAgentPingResultData(1, i), 6)
        Else
            itemx.SubItems(4) = ""
        End If
        
        'avg rtt
        If recvpkt > 0 Then
            itemx.SubItems(5) = Right(SPACE5 & Format(Round(aryAgentPingResultData(2, i) / recvpkt, 1), "###0.0"), 6) 'Avg RTT
        Else
            itemx.SubItems(5) = ""
        End If
        
        'max rtt
        If aryAgentPingResultData(3, i) >= 0 Then
            itemx.SubItems(6) = Right(SPACE5 & aryAgentPingResultData(3, i), 6)
        Else
            itemx.SubItems(6) = ""
        End If
        
        'sent
        itemx.SubItems(7) = glPingCount
        'received
        itemx.SubItems(8) = recvpkt
        'lost
        itemx.SubItems(9) = lostpkt
        'packet loss%
        itemx.SubItems(10) = Right(SPACE3 & Format(Round(100 * aryAgentPingResultData(5, i) / glPingCount, 1), "##0.0"), 6)
        'ping cycle(sec)
        itemx.SubItems(11) = Right(SPACE3 & Format(Round(aryPingStatData(6, i) / 1000, 1), "###0.0"), 6) 'elapsed time

        'continued fail
        If aryContinuedFail(i) > 0 Then
            itemx.SubItems(12) = Right(SPACE9 & aryContinuedFail(i), 10) 'continued fail
        Else
            itemx.SubItems(12) = "" '¥i¯à­«·sÂk¹s
        End If
        'alert count(§YUp --> Down)
        If aryPingStatData(5, i) > 0 Then
            itemx.SubItems(13) = Right(SPACE9 & aryPingStatData(5, i), 10) 'alert count
        End If
        
        'status code
        If aryAgentPingResultData(6, i) <> MY_PING_OK Then
            '***´ú¸Õ¥¢±Ñ
            itemx.SubItems(14) = aryAgentPingResultData(6, i) 'StatusCode GetStatusCode(StatusCode)
        Else
            itemx.SubItems(14) = "" 'StatusCode
        End If
  
        'Åã¥Ü§iÄµ
        Select Case aryPingStatData(1, i)
        Case RESULT_SUCCESS
        'Green
            If aryLastLedResult(i) <> RESULT_SUCCESS Then
                itemx.Text = " " '¸m©ó¤¤¶¡
                itemx.SmallIcon = "green"
                aryLastLedResult(i) = RESULT_SUCCESS
            End If
        Case RESULT_DOWN
        'Red
            If aryLastLedResult(i) <> RESULT_DOWN Then
                itemx.Text = "  "
                itemx.SmallIcon = "red"
                aryLastLedResult(i) = RESULT_DOWN
            End If
        Case RESULT_WARN
        'Yellow
            If aryLastLedResult(i) <> RESULT_WARN Then
                itemx.Text = "" '²¾¨ì³Ì«e­±,¦]¬°³Ì©_©Ç
                itemx.SmallIcon = "yellow"
                aryLastLedResult(i) = RESULT_WARN
            End If
        End Select
    Next
    
    'x = SendMessage(.hwnd, WM_SETREDRAW, 1, 0)
    End With
    Exit Sub
ErrHandler:
    MsgBox "RefreshPingStatus error!" & vbCrLf & Err.Description, vbExclamation, MsgTitle
End Sub

Public Sub AgentLedBlink(AgentID As Long, BlinkSwitch As Long)
    Dim idx As Long
    'aryAgentBlinkTick(AgentID) = GetTickCount
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

Public Sub AgentSayHello(AgentID As Long)
    Dim idx As Long
    idx = AgentID - 1
    picAgent(idx).Picture = picLed(3).Picture 'Say Hello®É§ï¦¨¶À¦â
    picAgent(idx).Refresh
End Sub

Public Sub AgentReady(AgentID As Long)
    Dim idx As Long
    idx = AgentID - 1
    picAgent(idx).Picture = picLed(1).Picture 'ºñ¦â
    picAgent(idx).Refresh
    '¥Î¥H¤Uªº¤èªk,¦]¬°¨ç¦¡¥i¯à¦P®É°õ¦æªº½t¬G,·|³y¦¨glAgentReadyCount¤@ª½³£¬O1
'    glAgentReadyCount = glAgentReadyCount + 1
'    '¥þ³¡ªºAgent¬Ò¤wReport Ready,§Ú¤]Report Ready
'    If glAgentReadyCount = glAgentCount Then
'        Call ReportReady
'    End If
End Sub
Public Sub AgentIsOff(AgentID As Integer)
    '°»´ú¤£¨ì,¤£¬OAgent¥D°Ê»¡GoodBye
    Dim idx As Long
    idx = AgentID - 1
    picAgent(idx).Picture = picLed(2).Picture '¬õ¦â
    picAgent(idx).Refresh
End Sub
Public Sub AgentSayGoodbye(AgentID As Long)
    Dim idx As Long
    idx = AgentID - 1
'    picAgent(idx).Visible = False
'    lblAgent(idx).Visible = False
    picAgent(idx).Picture = picLed(2).Picture '¬õ¦â
    picAgent(idx).Refresh
    aryAgentHwnd(AgentID) = 0
End Sub



Private Sub tmrPingQueue_Timer()
    Dim tick1 As Long
    Dim i As Long

    If PingQueueIndex > glAgentCount Then
        tmrPingQueue.Enabled = False
          '¤w±NtmrSummary©MRefreshList¨Ö¦b¤@°_
'            tmrSummary.Interval = glRefreshCycle ' glPingTimeOutBatch * 1.2
'            tmrSummary.Enabled = True
    Else
        If PingQueueIndex = 0 Then
            tmrPingQueue.Interval = glDelayStart
            TellAgentToDoSomething MSG_AGENT_STARTPING, glLogAgentHwnd
            tick1 = GetTickCount
            'glLogAgentTick = tick1
            PingQueueIndex = PingQueueIndex + 1
        Else
            TellAgentToDoSomething MSG_AGENT_STARTPING, aryAgentHwnd(PingQueueIndex)
            tick1 = GetTickCount
            
''                'ªì©l¤Ætick­È
            For i = aryAgentLB(PingQueueIndex) To aryAgentUB(PingQueueIndex)
                aryNodeLastStatisticsCycleTick(i) = tick1 '¥Î¨Ó­pºâping cycle interval
            Next
            PingQueueIndex = PingQueueIndex + 1
        End If
    End If

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
Public Sub MoveMyPos()
    Dim RCT As RECT
    Dim X As Variant, Y As Variant
    Dim ParentHwnd As Long
        ParentHwnd = GetParent(Me.hwnd)
        'MsgBox "ParentHwnd=" & ParentHwnd
        GetClientRect glPingManagerHwnd, RCT
        'MsgBox ParentHwnd
        'MsgBox Me.hwnd
'        X = (((RCT.Right - RCT.Left) * Screen.TwipsPerPixelX) - .Width) / 2
'        Y = (((RCT.Bottom - RCT.Top) * Screen.TwipsPerPixelY) - .Height) / 2
    On Error Resume Next
    Me.Move (RCT.Left + 2) * Screen.TwipsPerPixelX, (RCT.Top + 2) * Screen.TwipsPerPixelY, (RCT.Right - RCT.Left - 4) * Screen.TwipsPerPixelX, (RCT.Bottom - RCT.Top - 4) * Screen.TwipsPerPixelY


End Sub
Public Sub CloseMe()
    Unload Me
End Sub
Private Sub InitAgentLed()
    Dim i As Integer
    For i = 0 To MaxAgent - 1
        picAgent(i).Picture = picLed(0).Picture '¥þ³¡led³]¦¨¦Ç¦â
    Next
    'frLedPanel.Refresh
    Picture1.Refresh
End Sub
Public Sub InitAgentStuff()
    Dim i As Long
    
    ReDim aryAgentLB(1 To glAgentCount)
    ReDim aryAgentUB(1 To glAgentCount)
    ReDim aryAgentNumOfPingNode(1 To glAgentCount)
    ReDim aryAgentHwnd(1 To glAgentCount)
    ReDim aryAgentReady(1 To glAgentCount)
    ReDim aryAgentPingInfo(1 To 6, 1 To glAgentCount) 'pingcount, nodestart, nodeend, sn, ping cycle interval, step counter
    'ReDim aryAgentBlinkTick(1 To glAgentCount)
    ReDim aryAgentRefreshRange(1 To glAgentCount, 1 To 2) 'start, end

    Call InitAgentLed '¥þ³¡led³]¦¨¦Ç¦â
    For i = 0 To glAgentCount - 1
        lblAgent(i).Visible = True
        picAgent(i).Visible = True
    Next
    For i = glAgentCount To MaxAgent - 1
        lblAgent(i).Visible = False
        picAgent(i).Visible = False
    Next
    Picture1.Refresh
    
    '***­pºâ¨C­Óagentªºpinglistªºubound & lbound
    '***­pºâ¨C­ÓPingAgentªº¤W¤UPingNodeÃä¬É
    Dim num As Long
    Dim ModNum As Long
    Dim ub As Long, lb As Long
    
    
    num = NumOfPingNode \ glAgentCount
    ModNum = NumOfPingNode Mod glAgentCount

    ub = -1 'ªì©l¤Æ¤W¤@­Óub

    For i = 1 To glAgentCount
        lb = ub + 1
        If i <= ModNum Then
            ub = lb + num
        Else
            ub = lb + num - 1
        End If

        aryAgentLB(i) = lb
        aryAgentUB(i) = ub
        aryAgentNumOfPingNode(i) = ub - lb + 1
        '­n¥ýªì©l¤Ærefresh rangeªº±Ò©lÂI
        aryAgentRefreshRange(i, 1) = lb
        aryAgentRefreshRange(i, 2) = -1
'        MsgBox "i=" & i
'        MsgBox "lb=" & lb
'        MsgBox "ub=" & ub
    Next
    
End Sub
Public Sub ShowStatus(msg As String)
    statusbar.Panels(6).Text = msg
End Sub
