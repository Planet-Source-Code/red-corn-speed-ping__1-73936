VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Âù½u©T©w¹ï¸Ü¤è¶ô
   Caption         =   "About"
   ClientHeight    =   3915
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5790
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmAbout.frx":030A
   ScaleHeight     =   2702.202
   ScaleMode       =   0  '¨Ï¥ÎªÌ¦Û­q
   ScaleWidth      =   5437.109
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '¿Ã¹õ¤¤¥¡
   Begin VB.Timer tmrAutoClose 
      Left            =   5265
      Top             =   225
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   2550
      Left            =   90
      Picture         =   "frmAbout.frx":75A4C
      ScaleHeight     =   1748.81
      ScaleMode       =   0  '¨Ï¥ÎªÌ¦Û­q
      ScaleWidth      =   1843.625
      TabIndex        =   1
      Top             =   90
      Width           =   2685
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4245
      TabIndex        =   0
      Top             =   3405
      Width           =   1260
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '³z©ú
      Caption         =   "Jimmy Hung "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   2970
      TabIndex        =   5
      Top             =   2925
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  '¤º¹ê½u
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   2215.599
      Y2              =   2215.599
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  '³z©ú
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   2880
      TabIndex        =   2
      Top             =   1395
      Width           =   2850
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  '³z©ú
      Caption         =   "xyz"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   2970
      TabIndex        =   4
      Top             =   855
      Width           =   4245
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   2225.952
      Y2              =   2225.952
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  '³z©ú
      Caption         =   "The fast ping of the world!"
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   240
      TabIndex        =   3
      Top             =   3390
      Width           =   3735
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    'Me.Caption = "Ãö©ó " & App.Title
    lblTitle.Caption = MsgTitle
End Sub
Public Sub SetAutoClose(second As Long)
    tmrAutoClose.Interval = second * 1000
    tmrAutoClose.Enabled = True
End Sub

Private Sub tmrAutoClose_Timer()
    tmrAutoClose.Interval = 0
    tmrAutoClose.Enabled = False
    Unload Me
End Sub
