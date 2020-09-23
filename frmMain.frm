VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Network Change TCP/IP"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   855
   ClientWidth     =   8940
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   8940
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   8940
      TabIndex        =   15
      Top             =   5205
      Width           =   8940
      Begin VB.Line line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   10545
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   0
         X2              =   10545
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Label lblStatus 
         BackStyle       =   0  'Transparent
         Height          =   315
         Left            =   60
         TabIndex        =   16
         Top             =   30
         Width           =   8835
      End
   End
   Begin VB.PictureBox picTray 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   0
      Picture         =   "frmMain.frx":1982
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.CommandButton cmdEnableDHCP 
      Caption         =   "Attiva impostazioni automatiche"
      Height          =   600
      Left            =   5985
      TabIndex        =   6
      Top             =   4380
      Width           =   2775
   End
   Begin VB.CommandButton cmdProfileActivate 
      Caption         =   "Attiva il profilo selezionato"
      Enabled         =   0   'False
      Height          =   600
      Left            =   5985
      TabIndex        =   5
      Top             =   3675
      Width           =   2775
   End
   Begin VB.CommandButton cmdProfileEdit 
      Caption         =   "Modifica profilo..."
      Height          =   375
      Left            =   5985
      TabIndex        =   2
      Top             =   1620
      Width           =   2775
   End
   Begin VB.CommandButton cmdProfileDelete 
      Caption         =   "Elimina profilo"
      Height          =   390
      Left            =   5985
      TabIndex        =   4
      Top             =   2700
      Width           =   2775
   End
   Begin VB.CommandButton cmdProfileClone 
      Caption         =   "Clona profilo..."
      Enabled         =   0   'False
      Height          =   375
      Left            =   5985
      TabIndex        =   3
      Top             =   2160
      Width           =   2775
   End
   Begin VB.TextBox txtProfileSettings 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Left            =   300
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "frmMain.frx":2384
      Top             =   3675
      Width           =   5460
   End
   Begin VB.CommandButton cmdProfileCreateNew 
      Caption         =   "Crea un nuovo profilo..."
      Height          =   375
      Left            =   5985
      TabIndex        =   1
      Top             =   1080
      Width           =   2775
   End
   Begin VB.ListBox lstProfiles 
      Height          =   2010
      Left            =   300
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   1080
      Width           =   5445
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   7095
      TabIndex        =   14
      Top             =   720
      Width           =   615
   End
   Begin VB.Label lblActiveProfileName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   1290
      TabIndex        =   12
      Top             =   3120
      Width           =   3465
   End
   Begin VB.Label lblActiveProfile 
      BackStyle       =   0  'Transparent
      Caption         =   "Profilo attivo:"
      Height          =   195
      Left            =   300
      TabIndex        =   11
      Top             =   3120
      Width           =   990
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "lblTitle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   990
      TabIndex        =   10
      Top             =   330
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Impostazioni TCP/IP"
      Height          =   195
      Index           =   1
      Left            =   300
      TabIndex        =   9
      Top             =   3450
      Width           =   1455
   End
   Begin VB.Image imgVBCorner 
      Height          =   495
      Left            =   6945
      MouseIcon       =   "frmMain.frx":238A
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":24DC
      Stretch         =   -1  'True
      ToolTipText     =   "Visita www.vbcorner.net"
      Top             =   120
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seleziona il profilo desiderato dall'elenco"
      Height          =   195
      Index           =   0
      Left            =   300
      TabIndex        =   7
      Top             =   855
      Width           =   2835
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   300
      Picture         =   "frmMain.frx":2B74
      ToolTipText     =   "Credits"
      Top             =   210
      Width           =   480
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileMinimizeOnClickX 
         Caption         =   "Click on 'X' minimize window"
      End
      Begin VB.Menu mnuFileAutoSelectNetworkCard 
         Caption         =   "Auto-select network card (if one only exixts)"
      End
      Begin VB.Menu mnuFileInfo 
         Caption         =   "&Info..."
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuLanguages 
      Caption         =   "Language"
      Tag             =   "Multi-Language Support"
      Begin VB.Menu mnuLanguage 
         Caption         =   "mnuLanguage 0"
         Index           =   0
         Tag             =   "Multi-Language Support"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&?"
      Begin VB.Menu mnuHelpHomePage 
         Caption         =   "&Home page di vbCorner..."
      End
      Begin VB.Menu mnuHelpSendEmail 
         Caption         =   "Send e-mail to author..."
      End
      Begin VB.Menu mnuHelpSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About..."
      End
   End
   Begin VB.Menu mnuTray 
      Caption         =   "MenuTray"
      Begin VB.Menu mnuTrayOpen 
         Caption         =   "&Show window..."
      End
      Begin VB.Menu mnuTrayAbout 
         Caption         =   "&About..."
      End
      Begin VB.Menu mnuTraySep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTrayExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module      : frmMain
' DateTime    : 06/13/2005 16.02
' Author      : Giorgio Brausi (vbcorner@vbcorner.net - http://www.vbcorner.net)
' Project     : Network_Change
' Purpose     : Manage TCP/IP settings from Microsoft Visual Basic 6.0
' Descritpion :
' Comments    : last update 02/17/2005
'---------------------------------------------------------------------------------------
Option Explicit

Dim sComputer As String
Dim oWMIService As Object
Dim cNetAdapters As Object
Dim sIPAddress As Variant
Dim sSubnetMask As Variant
Dim sGateway As Variant
Dim sGatewaymetric As Variant
Dim sDNSServers As Variant
Dim sWINSPriServer As String
Dim sWINSSecServer As String
Dim oNetAdapter As Object
Dim bErr As Boolean

Dim IPConfigSet As Object
Dim IPConfig As Object

Const DHCP_NOT_ENABLED = 0
Const DHCP_ENABLED = -1
Const DHCP_CANCEL = -2

Rem Array that will contains profile settings
Dim sAr() As String

Dim bIsLoading As Boolean

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If bIsIde Then Exit Sub
  Select Case UnloadMode
    Case vbFormControlMenu
      If mnuFileMinimizeOnClickX.Checked = True Then
        Cancel = True
        Me.WindowState = vbMinimized
      End If
  End Select
End Sub

Private Sub mnuFileAutoSelectNetworkCard_Click()
  Rem if this options is activated, when you activate a profile the network card
  Rem will be selected automatically (if one only network card exixst)
  
  mnuFileAutoSelectNetworkCard.Checked = Not mnuFileAutoSelectNetworkCard.Checked
  gbAutoSelectIfOnlyCardExists = mnuFileAutoSelectNetworkCard.Checked
  
  Rem Save
  AppINI.WriteString "Options", "AutoSelectIfOnlyONECardExists", gbAutoSelectIfOnlyCardExists
  
  
End Sub

Private Sub mnuFileMinimizeOnClickX_Click()
  mnuFileMinimizeOnClickX.Checked = Not mnuFileMinimizeOnClickX.Checked
  
  giMinimizeOnClickX = mnuFileMinimizeOnClickX.Checked
  
End Sub

Private Sub mnuHelpAbout_Click()
  Rem Show informations about the program by using the
  Rem current language id.
  frmAbout.Show vbModal, Me
  
  Set frmAbout = Nothing
  
End Sub

Private Sub mnuHelpHomePage_Click()
  OpenUrl URL_WEB_SITE
End Sub

Private Sub mnuHelpSendEmail_Click()
  Dim sEmailAddress As String, sSubject As String, sBodyText As String
  
  sEmailAddress = "vbcorner@vbcorner.net"
  sSubject = "Network Change " & App.Major & App.Minor & App.Revision
  sBodyText = "Hi, Giorgio" & gsCR & "I contact you about your project:" & gsCR & gsCR & sSubject & gsCR & gsCR & "I want know..." & gsCR & gsCR & gsCR & gsCR
  
  SendEMail Me.hWnd, sEmailAddress, sSubject, sBodyText
  
End Sub

Private Sub mnuTrayAbout_Click()
  mnuHelpAbout_Click
End Sub

Private Sub mnuTrayExit_Click()
  Unload Me
End Sub

Private Sub mnuTrayOpen_Click()

  On Error Resume Next
  ActivateForm Me
  DestroyIcon
  
  
End Sub





Private Sub cmdProfileActivate_Click()
  Dim sMsg As String
  
  If lstProfiles.ListIndex = -1 Then
    Rem Select a profile, first!
    sMsg = MLSGetString("0002") ' MLS-> "Selezionare prima un profilo dalla lista!"
    MsgBox sMsg, vbInformation
    Exit Sub
  Else
    Rem Confirm request
    sMsg = MLSGetString("0025") ' MLS-> "Attivare il profilo " & lstProfiles.Text & " ?"
    sMsg = Replace(sMsg, "%s", lstProfiles.Text)
    If MsgBox(sMsg, vbQuestion + vbYesNo) = vbNo Then Exit Sub
  End If
  
  Rem Select network card (NIC)
  If EachNIC = False Then
    sMsg = MLSGetString("0040") ' MLS-> "Nessun Profilo Attivato"
    MsgBox sMsg, vbInformation
    Exit Sub
  End If
  
  lblStatus.Caption = MLSGetString("0031")
  
  MsgBox lstProfiles.Text & " " & MLSGetString("0028")
  lblActiveProfileName.Caption = lstProfiles.Text
  
  Rem Register profile on INI file
  CIni.WriteString "CurrentProfile", "ProfileName", lstProfiles.Text
  
  Rem Update form caption
  Dim sProfile As String
  sProfile = CIni.ReadString("CurrentProfile", "ProfileName", "")
  Me.Caption = App.ProductName & IIf(sProfile <> "", " (" & sProfile & ")", "")
  
End Sub

Private Sub cmdProfileActivate_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblStatus.Caption = MLSGetString("0003") ' MLS-> "Applica l'impostazione TCP/IP selezionata"
End Sub


Private Sub cmdProfileClone_Click()

  eProfileMode = PROF_MODE_SAVE
  
  ProfileEdit
  
End Sub

Private Sub cmdProfileClone_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblStatus.Caption = MLSGetString("0004") ' MLS-> "Crea un nuovo profilo partendo dal profilo esistente selezionato"
End Sub


'---------------------------------------------------------------------------------------
' Procedure   : cmdProfileCreateNew_Click
' DateTime    : 12/06/2005 17.38
' Author      : Giorgio Brausi (vbcorner@vbcorner.net - http://www.vbcorner.net)
' Purpose     : Open form to create new profiles
' Descritpion : The new profile created will append immediately on the profile list
'             :
'---------------------------------------------------------------------------------------
Private Sub cmdProfileCreateNew_Click()

  eProfileMode = PROF_MODE_SAVE
  
  Rem Open a new instace of form
  Dim frmAddNew As frmNewSettings
  Set frmAddNew = New frmNewSettings
  frmAddNew.Show vbModal, Me
  Set frmAddNew = Nothing
  
End Sub

Private Sub cmdProfileCreateNew_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblStatus.Caption = MLSGetString("0005") ' MLS-> "Crea un nuovo profilo da aggiungere alla lista"
End Sub


Private Sub cmdEnableDHCP_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblStatus.Caption = MLSGetString("0006") ' MLS-> "Azzera le impostazioni TCP/IP e le imposta su 'Ottieni automaticamente indirizzi IP e DNS...'"
End Sub

Private Sub cmdEnableDHCP_Click()
  Dim sMsg As String
  Dim lRet As Long
  
  Select Case SetTCP_Automatic
    Case DHCP_ENABLED
      sMsg = MLSGetString("0007") ' MLS-> "DHCP è stato abilitato."
    Case DHCP_NOT_ENABLED
      sMsg = MLSGetString("0008") ' MLS-> "DHCP non è stato abilitato."
    Case DHCP_CANCEL
      Exit Sub
  End Select
  lblActiveProfileName.Caption = ""
  
  Rem Update INI file
  CIni.WriteString "CurrentProfile", "ProfileName", ""
   
  Rem Update application title
  Dim sProfile As String
  sProfile = CIni.ReadString("CurrentProfile", "ProfileName", "")
  Me.Caption = App.ProductName & IIf(sProfile <> "", " (" & sProfile & ")", "")
  
  Rem Disable profile and settings
  lstProfiles.ListIndex = -1
  txtProfileSettings.Text = ""
  cmdProfileActivate.Enabled = False
  
  MsgBox sMsg, vbInformation
  
End Sub


Private Sub cmdProfileDelete_Click()
  Dim sMsg As String
  
  If lstProfiles.ListIndex = -1 Then
    sMsg = MLSGetString("0009") ' MLS-> "Selezionare il profilo da eliminare dall'elenco!"
    MsgBox sMsg, vbCritical
    Exit Sub
  Else
    sMsg = MLSGetString("0026") ' MLS-> "Eliminare il profilo " & lstProfiles.Text & " ?"
    sMsg = Replace(sMsg, "%s", lstProfiles.Text)
    If MsgBox(sMsg, vbCritical + vbYesNo) = vbNo Then Exit Sub
  End If

  txtProfileSettings.Text = ""
  
  CIni.DeleteKey INI_SECTION_PROFILES, lstProfiles.Text
  
  LoadProfiles
  
  
End Sub

Private Sub cmdProfileDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblStatus.Caption = MLSGetString("0010") ' MLS-> "Elimina il profilo selezionato."
End Sub


Private Sub cmdProfileEdit_Click()
  eProfileMode = PROF_MODE_EDIT
  ProfileEdit
End Sub

Private Sub cmdProfileEdit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblStatus.Caption = MLSGetString("0011") ' MLS-> "Modifica il profilo selezionato."
End Sub


Private Sub Form_Deactivate()
  lblStatus.Caption = ""
End Sub

Private Sub Form_Load()

  bIsLoading = True
  
  MLSFillMenuLanguages '<- Add by Multi-Languages Support Add-in
  MLSLoadLanguage Me   '<- Add by Multi-Languages Support Add-in
 
  mnuTray.Visible = False
  
  App.Title = frmMain.Caption
  lblTitle = App.ProductName
  
  Set CIni = New clsIni
  CIni.FileName = App.Path & "\" & "PROFILES.INI"
  
  Set AppINI = New clsIni
  
  txtProfileSettings.Text = ""
  
  LoadSettings
  LoadProfiles
  
  lblVersion.Caption = MLSGetProperty(Me.Name, "lblVersion.Caption") & " " & App.Major & "." & App.Minor & "." & App.Revision
  'lblVersion.ForeColor = RGB(41, 154, 206)
  
  With imgIcon
    .MousePointer = imgVBCorner.MousePointer
    .MouseIcon = imgVBCorner.MouseIcon
  End With
  
  Rem No text will selected
  'bOnFocus = False
  
  Rem Update application title
  Dim sProfile As String
  sProfile = CIni.ReadString("CurrentProfile", "ProfileName", "")
  Me.Caption = App.ProductName & IIf(sProfile <> "", " (" & sProfile & ")", "")
  If sProfile <> "" Then
    lblActiveProfileName.Caption = sProfile
    lstProfiles.Text = sProfile
  Else
    lblActiveProfileName.Caption = "(none)"
  End If
  
  Rem If you put a shortcut on Startup (Start menu) you can hide Network Change
  Rem by using the '/AUTO' command parameter
  If UCase(Command$) = "/HIDEONSTARTUP" Or UCase(Command$) = "HIDEONSTARTUP" Then
    Me.WindowState = vbMinimized
  End If
  
  
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  
  If X > cmdProfileClone.Left And X < (cmdProfileClone.Left + cmdProfileClone.Width) And Y > cmdProfileClone.Top And Y < (cmdProfileClone.Top + cmdProfileClone.Height) Then
    lblStatus.Caption = MLSGetString("0012") ' MLS-> "Crea un nuovo profilo partendo dal profilo esistente selezionato"
  ElseIf lblStatus.Caption <> "" Then
    lblStatus.Caption = ""
  End If
End Sub


Private Sub Form_Resize()
  If Me.WindowState = vbMinimized Then
    Me.Hide
    InitTrayIcon picTray, picTray.Picture, Me.Caption
  Else
    Dim sProfile As String
    sProfile = CIni.ReadString("CurrentProfile", "ProfileName", "")
    Me.Caption = App.ProductName & IIf(sProfile <> "", " (" & sProfile & ")", "")
    If gbIconInTray Then
      DestroyIcon
    End If
  
  End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  Set CIni = Nothing
  DestroyIcon
  DoEvents
End Sub


Private Sub imgIcon_Click()
  Dim s As String
  
  s = MLSGetString("0013") & vbCrLf ' MLS-> "Un ringraziamento a:"
  s = s & "- Mario Raccagni (mariora@aznet.it)" & vbCrLf
  s = s & "- Roberto Doretto (roberto.doretto@volvo.com)" & vbCrLf
  
  MsgBox s, vbExclamation
  
End Sub

Private Sub imgVBCorner_Click()
  OpenUrl "http://www.vbcorner.net/eng"
End Sub

Private Sub lstProfiles_Click()
  Dim sSettings As String, i As Integer
  Dim sIP As String, sSM As String, sGW As String
  Dim sDNS1 As String, sDNS2 As String
  
  Rem Enable Clone button
  cmdProfileClone.Enabled = True
  cmdProfileActivate.Enabled = True
  
  Rem Read profile settings
  sSettings = CIni.ReadString(INI_SECTION_PROFILES, lstProfiles.Text, "")
  sAr = Split(sSettings, ".", , vbTextCompare)
  
  If UBound(sAr) < 1 Then
    txtProfileSettings.Text = ""
    Exit Sub
  End If
  
  Rem Show profile settings, in according to language:
    sIP = MLSGetString("0032") & vbTab & sAr(0) & "," & sAr(1) & "," & sAr(2) & "," & sAr(3)
    sSM = MLSGetString("0033") & vbTab & sAr(4) & "," & sAr(5) & "," & sAr(6) & "," & sAr(7)
    sGW = MLSGetString("0034") & vbTab & sAr(8) & "," & sAr(9) & "," & sAr(10) & "," & sAr(11)
  sDNS1 = MLSGetString("0035") & vbTab & sAr(12) & "," & sAr(13) & "," & sAr(14) & "," & sAr(15)
  sDNS2 = MLSGetString("0036") & vbTab & sAr(16) & "," & sAr(17) & "," & sAr(18) & "," & sAr(19)
  txtProfileSettings.Text = sIP & vbCrLf & sSM & vbCrLf & sGW & vbCrLf & sDNS1 & vbCrLf & sDNS2

End Sub



'---------------------------------------------------------------------------------------
' Procedure   : RetAddr
' DateTime    : 12/06/2005 17.25
' Author      : Giorgio Brausi (vbcorner@vbcorner.net - http://www.vbcorner.net)
' Purpose     : Return a formatted string like: "xxx.xxx.xxx.xxx"
' Descritpion : This fucntion accept an array (sAr) and the first array index
'             : (idx) from where start. Example:
'             :   sIPAddress = Array(RetAddr(sAr(), 0))
'             : Will be formatted the string with the firsti 4 array (from 0 to 3)
'             : If the array contains this next values:
'             : sAr(0) = 192, sAr(1) = 168, sAr(2) = 1 e sAr(0) = 2
'             : RetAddr return the string: "192.168.1.2"
'---------------------------------------------------------------------------------------
Public Function RetAddr(ByRef ar As Variant, idx As Integer) As String
  
  Dim k As Integer, sTmp As String
  
  For k = idx To idx + 3
    Rem Format the stringa, add the dot (.)
    sTmp = sTmp & "." & ar(k)
  Next k
  
  Rem Return address, first dot is removed.
  RetAddr = Mid(sTmp, 2)
  
End Function

'---------------------------------------------------------------------------------------
' Procedure   : SetTCP_Profile
' DateTime    : 12/06/2005 17.35
' Author      : Giorgio Brausi (vbcorner@vbcorner.net - http://www.vbcorner.net)
' Purpose     : Active selected profile
' Descritpion : When a profile is selected, the array sAr() will filled
'             :
' Used by     : Calling from cmdProfileActivate_Click
'---------------------------------------------------------------------------------------
'
Public Sub SetTCP_Profile(NetworkCards As String)
  
  sComputer = "."  ' Computer Name (.) for local computer
    
  Set oWMIService = GetObject("winmgmts:\\" & sComputer & "\root\cimv2")
  Set cNetAdapters = oWMIService.ExecQuery("Select * FROM Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")

  Rem ---------------------------------------------------------------------------
  Rem Sample for TCP/IP parameters
  '  sIPAddress = Array("192.168.1.2")                  ' IP
  '  sSubnetMask = Array("255.0.0.0")                   ' Mask
  '  sGateway = Array("192.168.1.1")                    ' Gateway
  '  sGatewaymetric = Array(1)                          ' Metrica
  '  sDNSServers = Array("192.168.1.3", "192.168.1.4")  ' DNS, also multiple
  '  sWINSPriServer = "192.168.1.5"                     ' WINS primary
  '  sWINSSecServer = "192.168.1.6"                     ' WINS secondary
  Rem ---------------------------------------------------------------------------
  
  sIPAddress = Array(RetAddr(sAr(), 0))
  sSubnetMask = Array(RetAddr(sAr(), 4))
  sGateway = Array(RetAddr(sAr(), 8))
  sGatewaymetric = Array(1)
  sDNSServers = Array(RetAddr(sAr(), 12), RetAddr(sAr(), 16))
  sWINSPriServer = "0.0.0.0"
  sWINSSecServer = "0.0.0.0"
  'sWINSPriServer = "192.168.1.5"  ' <- Set if need
  'sWINSSecServer = "192.168.1.6"  ' <- Set if need
  
  For Each oNetAdapter In cNetAdapters
      Rem If NetWorkCard = *ALL the profile will activated
      If InStr(1, oNetAdapter.Caption, NetworkCards) > 0 Or NetworkCards = "*ALL" Then
            bErr = oNetAdapter.EnableStatic(sIPAddress, sSubnetMask)
            bErr = oNetAdapter.SetGateways(sGateway, sGatewaymetric)
            bErr = oNetAdapter.SetDNSServerSearchOrder(sDNSServers)
            bErr = oNetAdapter.SetWINSServer(sWINSPriServer, sWINSSecServer)
      End If
  Next

End Sub

Private Sub lstProfiles_DblClick()
  cmdProfileEdit.Value = True
End Sub

Private Sub lstProfiles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblStatus.Caption = MLSGetString("0014") ' MLS-> "Clic: seleziona il profilo da attivare - Doppio Clic: modifica il profilo."
End Sub


Private Sub mnuFileExit_Click()
  Unload Me
End Sub

Private Sub mnuFileInfo_Click()
  Call imgIcon_Click
End Sub

Private Sub picTray_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Rem This routine receive the messages form tray icon
    Dim Msg As Long
    Rem X value change in according to ScaleMode setting.
    Msg = picTray.ScaleX(X, picTray.ScaleMode, vbPixels)
       
    Select Case Msg
        Case WM_LBUTTONUP
          Rem LEFT button UP
          'ActivateForm Me
          'DestroyIcon '/ l'icona non serve più

        Case WM_LBUTTONDBLCLK
            Rem LEFT double-click
            ActivateForm Me ' Show form
            DestroyIcon     ' Icon don't longer need
            
        Case WM_RBUTTONUP
          Rem RIGHT Click button
          PopupMenu mnuTray ' Show popup menu
          
    End Select

End Sub

Private Sub txtProfileSettings_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblStatus.Caption = MLSGetString("0015") ' MLS-> "Impostazioni del profilo selezionato."
End Sub



'---------------------------------------------------------------------------------------
' Procedure   : SetTCP_Automatic
' DateTime    : 16/06/2005 01.31
' Author      : Giorgio Brausi (vbcorner@vbcorner.net - http://www.vbcorner.net)
' Purpose     : Reset all TCP/IP settings and say to computer to obtain
'             : IP and DNS address automatically
' Comment     : if [bShowConfirm] = True, a confirm message is displayed.
'---------------------------------------------------------------------------------------
Private Function SetTCP_Automatic(Optional bShowConfirm As Boolean = True) As Long
'**************************************************************
'     SetDHCP.vbs
'**************************************************************
Dim wmiLocator As Object
Dim wmiService As Object
Dim colNetAdapters As Object
Dim objNetAdapter As Object
Dim strIPAddress As Variant
Dim strSubnetMask As Variant
Dim strGateway As Variant
Dim strGatewayMetric As Variant
Dim strDNSServers As Variant
Dim errEnable As Variant
Dim errGateways As Variant
Dim errDNSServers As Variant

Dim sMsg As String

  If bShowConfirm Then  ' If
      Rem Request confirm...
      sMsg = MLSGetString("0016") & vbCrLf & MLSGetString("0017") ' MLS-> "Attenzione: saranno azzerati tutti i parametri TCP/IP." ' MLS-> "Confermate?"
      
      If MsgBox(sMsg, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
        SetTCP_Automatic = DHCP_CANCEL
        Exit Function
      End If
  End If
  
  Set wmiLocator = CreateObject("WbemScripting.SWbemLocator")
  Set wmiService = wmiLocator.ConnectServer("localhost", "root\cimv2")
  wmiService.Security_.ImpersonationLevel = 3
  
  Set colNetAdapters = wmiService.ExecQuery _
  ("Select * from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")
  
  'On Error Resume Next
  
  For Each objNetAdapter In colNetAdapters
    errDNSServers = objNetAdapter.SetDNSServerSearchOrder()
    errEnable = objNetAdapter.EnableDHCP()
    If errEnable = 0 Then
      SetTCP_Automatic = DHCP_ENABLED       'DHCP has been enabled.
    Else
      SetTCP_Automatic = DHCP_NOT_ENABLED   'DHCP could not be enabled.
    End If
  Next

End Function

'---------------------------------------------------------------------------------------
' Procedure   : LoadProfiles
' DateTime    : 16/06/2005 14.33
' Author      : Giorgio Brausi (vbcorner@vbcorner.net - http://www.vbcorner.net)
' Purpose     : Load all profile to ListBox
'---------------------------------------------------------------------------------------
Private Sub LoadProfiles()
  ProfilesLoad CIni.FileName, lstProfiles
End Sub

'---------------------------------------------------------------------------------------
' Procedure   : ProfileEdit
' DateTime    : 16/06/2005 14.32
' Author      : Giorgio Brausi (vbcorner@vbcorner.net - http://www.vbcorner.net)
' Purpose     : This routine edit or clone a profile
'
'---------------------------------------------------------------------------------------
Public Sub ProfileEdit()
  Dim sMsg As String, i As Integer
  Const STR_CLONE = " (CLONE)"
  
  If lstProfiles.ListCount = 0 Then
    sMsg = MLSGetString("0018") ' MLS-> "Non esiste alcun profilo!"
  ElseIf lstProfiles.ListIndex = -1 Then
    sMsg = MLSGetString("0019") ' MLS-> "Selezionare prima un profilo dalla lista!"
  End If
  If sMsg <> "" Then
    MsgBox sMsg, vbInformation
    Exit Sub
  End If
  
  Rem Open a new instance of form to modify/clone a profile
  Dim frmAddNew As frmNewSettings
  Set frmAddNew = New frmNewSettings
  Load frmAddNew
  

  With frmAddNew
    Rem Set the form title in according to action
    If eProfileMode = PROF_MODE_EDIT Then
      .Caption = MLSGetString("0020") ' MLS-> "Modifica Profilo"
    ElseIf eProfileMode = PROF_MODE_SAVE Then
      .Caption = MLSGetString("0021") ' MLS-> "Clona Profilo"
    End If
    
    Rem Load all profile values
    For i = 0 To .txtIP.UBound
      .txtIP(i) = sAr(i)
    Next i
    
    If eProfileMode = PROF_MODE_EDIT Then
      Rem Modify a existing profile
      .txtProfileName = lstProfiles.Text
    Else
      Rem Clone a existing profile
      .txtProfileName = lstProfiles.Text & STR_CLONE
    End If
    .Show vbModal, Me
  End With
  
  Rem Release form instance
  Set frmAddNew = Nothing

End Sub

Public Sub MLSFillMenuLanguages()
Rem This sub is created by Multi-Language Support
Rem Search for all language file, load & fill the menu item array
Rem named: mnuLanguage
Dim sMsg As String
    Dim iNumLanguages As Integer, i As Integer, sFileName As String

    Rem Search for language files. Folder is App.Path
    Rem Keep the LNG files on the App.Path folder!
    sFileName = Dir(App.Path & "\*.lng")
    If sFileName = "" Then
        sMsg = MLSGetString("0027") ' MLS -> "Non è presente alcun file di linguaggio!"
        MsgBox sMsg
        Exit Sub
    End If

    Rem Now, for each language file add a new menu item (entry)
    Rem -----------------------------------------------
    Do While sFileName <> ""
        If iNumLanguages > 0 Then
            Load mnuLanguage(iNumLanguages)
            mnuLanguage(iNumLanguages).Visible = True
        End If
        mnuLanguage(iNumLanguages).Caption = Mid(sFileName, 1, Len(sFileName) - 4)
        iNumLanguages = iNumLanguages + 1
        sFileName = Dir
    Loop

    Rem Get the current language from the file "LangSetting.ini"
    Rem Note: the file LangSetting.ini is create by MLS, but if you want
    Rem you can delete it and add this section in your custom INI file, if there.
    If iNumLanguages > 0 Then
        gsLanguageFile = MLSReadINI(App.Path & "\" & "LangSetting.ini", "Language", "CurrentLanguage")
    End If
    For i = 0 To iNumLanguages - 1
        If mnuLanguage(i).Caption = gsLanguageFile Then
            mnuLanguage(i).Checked = True
            Exit For
        End If
    Next i

End Sub

Private Sub mnuLanguage_Click(Index As Integer)
    Dim i As Integer

    For i = 0 To mnuLanguage.UBound
        mnuLanguage(i).Checked = False
    Next i

    Rem Set the selected language
    mnuLanguage(Index).Checked = True
    gsLanguageFile = mnuLanguage(Index).Caption

    Rem Update all loaded forms for new language
    For i = 0 To Forms.Count - 1
        MLSLoadLanguage Forms(i)
    Next i

    Rem Update application title
    App.Title = frmMain.Caption
    lblTitle = App.Title

    Rem Update the CurrentLanguage entry to LangSetting.ini
    MLSWriteINI App.Path & "\LangSetting.ini", "Language", "CurrentLanguage", mnuLanguage(Index).Caption

    lblVersion.Caption = MLSGetProperty(Me.Name, "lblVersion.Caption") & " " & App.Major & "." & App.Minor & "." & App.Revision

    Rem If a profile is selected, update
    Rem otherwise empty TextBox
    If lstProfiles.ListIndex = -1 Then
        txtProfileSettings.Text = ""
    Else
        lstProfiles_Click
    End If
    
End Sub

Public Sub LoadSettings()
  
  With AppINI
    .FileName = gsAppPath & "CONFIG.INI"
    giMinimizeOnClickX = .ReadInt("Options", "MinimizeOnClickX", vbChecked)
    mnuFileMinimizeOnClickX.Checked = giMinimizeOnClickX
    gbAutoSelectIfOnlyCardExists = .ReadBool("Options", "AutoSelectIfOnlyONECardExists", False)
  End With
  
End Sub

Private Function EachNIC() As Boolean
    Dim frmNIC As frmSelectNIC
    Dim NetWorkCard As String
    
    Rem Call the form to select the network card where profile will be activated.
    Rem If property SelectedNIC = *ALL the profile will be activated for ALL network
    Rem cards, otherwise will be activated for selected network card only.
    Rem Note that if the option
    
    Rem Load a new instance of the form
    Set frmNIC = New frmSelectNIC
    Load frmNIC
    
    frmNIC.Show vbModal, Me
    
    Screen.MousePointer = vbHourglass
    lblStatus.Caption = MLSGetString("0030")
    DoEvents
    NetWorkCard = frmNIC.SelectedNIC
    
    Rem Relelase form object
    Set frmNIC = Nothing
    
    If Trim$(NetWorkCard & "") = "" Then
        EachNIC = False
    Else
        Rem Version 1.3.2: BEFORE to set the new TCP settings
        Rem RESET to 0 all currente settings, so I'm sure that the
        Rem new settings is applied correctly.
        If SetTCP_Automatic(False) = DHCP_ENABLED Then   ' False=no message
          SetTCP_Profile NetWorkCard
          EachNIC = True
        End If
    End If
    
    Screen.MousePointer = vbNormal
End Function
