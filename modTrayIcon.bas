Attribute VB_Name = "modTrayIcon"
Option Explicit
Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long

Rem ==================================================
Rem Module to manage tray icon
Rem To use this code you must add a PictureBox that
Rem contains the icon to show on tray-area.
Rem NOTE: Name it "IconInTray" !
Rem ==================================================
Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200

Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_RBUTTONUP = &H205
Public Const WM_LBUTTONUP As Long = &H202

Public Declare Function EnumWindows Lib "user32.dll" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function OpenIcon Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

Public NID As NOTIFYICONDATA

Dim m_sPattern As String
Dim m_lhFind As Long

Public gbFromIcon As Boolean
Public gi_nIco As Integer
Public gbIconInTray As Boolean

Sub DestroyIcon()
  On Error Resume Next
  Rem If the icon no longer need, destroy it
  Call Shell_NotifyIcon(NIM_DELETE, NID)
  gbIconInTray = False
End Sub
Public Sub InitTrayIcon(callback As Object, icon As IPictureDisp, testo As String)

  Rem The Form MUST be visible before to call Shell_NotifyIcon
  With NID
    .cbSize = Len(NID)
    .hwnd = callback.hwnd
    .uID = vbNull
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    .uCallbackMessage = WM_MOUSEMOVE
    .hIcon = icon.Handle
    .szTip = testo & vbNullChar
  End With
  gbIconInTray = True
  Shell_NotifyIcon NIM_ADD, NID

End Sub


'---------------------------------------------------------------------------------------
' Procedure   : FindPrevInstance
' DateTime    : 28/10/2004 13.12
' Author      : Giorgio Brausi (vbcorner@vbcorner.net - http://www.vbcorner.net)
' Purpose     : Serach is a previous instance is already running
' Descritpion : This routine will be execute 'always' on program
'             : start.
'             :
' Comments    : The PrevInstance property of App object is good but
'             : there is a problem: if one of the instances is run
'             : from VB IDE, this instance isn't recognized by
'             : PrevInstance property. thus, you can't to test your
'             : source code, you must to compile and test the EXE!
'             : With this routine, you find the previous instance
'             : even is this instace is run by VB IDE. This will
'             : be able to perform Debug actions!
'---------------------------------------------------------------------------------------
Public Function FindPrevInstance(ByVal sTitle As String) As Boolean
    Dim lhWnd As Long
    
    Rem Search window by Title
    Rem Cercare la finestra in base al titolo
    lhWnd = FindWindowHwnd(sTitle)
        
    If lhWnd > 0 Then
        Rem If found it restore the 1st instance
        Rem Se la trova recupera al 1ª instanza
        FindPrevInstance = True
        
        Rem Restore window state
        Rem Ripristina lo stato della finestra
        OpenIcon lhWnd
        
        Rem Bring to up
        Rem La porto in primo piano
        SetForegroundWindow lhWnd
        
    End If
    
End Function

Public Function EnumWinProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
    Dim k As Long
    Dim sName As String
    Rem fill the buffer
    Rem Riempie il buffer
    sName = Space$(128)
    Rem Get the window title
    Rem Rileva il titolo della finestra
    k = GetWindowText(hwnd, sName, 128)
    If k > 0 Then
        Rem Trim the caption
        Rem Rimuove gli spazi inutili
        sName = Left$(sName, k)
        If lParam = 0 Then sName = UCase(sName)
        Rem If window is found
        Rem Se la finestra viene trovata
        If sName Like m_sPattern Then
            Rem Return the window handle
            Rem Restituisce l'handle della finestra
            m_lhFind = hwnd
            Rem Then exit from function
            Rem quindi esce dalla funzione
            EnumWinProc = 0
            Exit Function
        End If
    End If
    EnumWinProc = 1
End Function
Public Function FindWindowHwnd(sWild As String) As Long
  Rem Save the window title
  Rem Salva il titolo della finestra
  m_sPattern = UCase$(sWild)
  
  Rem Enumerate all the opened windows
  Rem Enumera tutte le finestre aperte
  EnumWindows AddressOf EnumWinProc, False
  
  Rem Return the window handle
  Rem Restituisce l'handle della finestra
  FindWindowHwnd = m_lhFind
  
End Function

Public Sub ActivateForm(F As Form)
  Rem Activate the form and bring it up
  Rem Attiva il form e lo porta in primo piano
  F.WindowState = vbNormal
  SetForegroundWindow F.hwnd
  F.Show
  
End Sub

