VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informazioni su MiaApplicazione"
   ClientHeight    =   2565
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5040
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770.409
   ScaleMode       =   0  'User
   ScaleWidth      =   4732.821
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   3510
      TabIndex        =   0
      Top             =   2070
      Width           =   1260
   End
   Begin VB.Image imgHome 
      Height          =   240
      Left            =   540
      Picture         =   "frmAbout.frx":058A
      Top             =   2040
      Width           =   240
   End
   Begin VB.Label lblUrl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "www.vbcorner.net"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1050
      MouseIcon       =   "frmAbout.frx":0B14
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2085
      Width           =   1305
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   285
      Picture         =   "frmAbout.frx":0C66
      ToolTipText     =   "Credits"
      Top             =   195
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   42.257
      X2              =   4499.936
      Y1              =   1304.512
      Y2              =   1304.512
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lblDescription"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1050
      TabIndex        =   1
      Top             =   945
      Width           =   3420
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Network Change TCP/IP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   1050
      TabIndex        =   2
      Top             =   240
      Width           =   3390
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   42.257
      X2              =   4499.936
      Y1              =   1314.865
      Y2              =   1314.865
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1050
      TabIndex        =   3
      Top             =   585
      Width           =   645
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

    Dim sAbout As String, s As String
    s = MLSGetProperty(frmMain.Name, "mnuHelpAbout.Caption")
    s = Replace(s, "&", "")
    sAbout = Replace(s, ".", "")
    Me.Caption = sAbout & " Network Change TCP/IP"
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
  
    s = MLSGetString("0029")
    Me.lblDescription = s & vbCrLf & vbCrLf & "by " & App.LegalCopyright
   
    s = MLSGetProperty(frmMain.Name, "mnuHelpHomePage.caption")
    s = Replace(s, "&", "")
    s = Replace(s, ".", "")
    
    lblUrl.ToolTipText = s
    With imgHome
      .MouseIcon = lblUrl.MouseIcon
      .ToolTipText = s
      .MousePointer = vbCustom
    End With
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblUrl.Font.Underline = False
End Sub


Private Sub lblUrl_Click()
  OpenUrl URL_WEB_SITE
End Sub


Private Sub lblUrl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblUrl.Font.Underline = True
End Sub


