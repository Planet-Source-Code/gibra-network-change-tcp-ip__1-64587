Attribute VB_Name = "modMain"
Option Explicit

Public Const URL_WEB_SITE = "http://www.vbcorner.net"

Public gsAppPath As String
Public AppINI As clsIni

Public giMinimizeOnClickX As Integer
Public gbAutoSelectIfOnlyCardExists As Boolean

Public bIsIde As Boolean

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Const COLOR_BTNFACE = 15 'Button
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Const FLOODFILLBORDER = 0  ' Fill until crColor& color encountered.
Const FLOODFILLSURFACE = 1 ' Fill surface until crColor& color not encountered.
Dim mBrush As Long

Public Function IsIDE() As Boolean
    On Error GoTo Out
    Debug.Print 1 / 0
Out:
    IsIDE = Err
End Function

Public Function ValidateNumberChar(valAscii As Integer) As Integer
  Rem Allow only number (0-9) or dot '.'
  If valAscii < 32 Then
    ValidateNumberChar = valAscii
    Exit Function
  End If
  ValidateNumberChar = IIf(InStr(1, "1234567890.", Chr$(valAscii)) > 0, valAscii, 0)
End Function



'---------------------------------------------------------------------------------------
' Procedure   : OnFocus
' DateTime    : 08/02/2006 02.20
' Author      : Giorgio Brausi
' Purpose     : Change backcolor and select text
' Descritpion :
' Comments    :
'---------------------------------------------------------------------------------------
Public Sub OnFocus(Obj As Control, bState As Boolean)

  If bOnFocus = False Then Exit Sub
  
  On Error Resume Next
  If bState = True Then
    Obj.BackColor = &HC0FFFF
    Obj.SelStart = 0
    Obj.SelLength = 1000
  Else
    Obj.BackColor = vbWindowBackground
  End If
End Sub


Public Sub Main()
  Rem Use Windows XP Style (THEMES)
  Rem "Network_Change.exe.manifest" file is request!
  InitCommonControls
  
  gsAppPath = SetPath()
  
  bIsIde = IsIDE()
  
  Rem Show main form
  frmMain.Show
  
End Sub


Public Function SetPath() As String

  Dim sVarPath As String
  
  sVarPath = App.Path
  If Right(sVarPath, 1) <> "\" Then
    sVarPath = sVarPath & "\"
  End If
  
  SetPath = sVarPath
  
End Function

Public Function FloodFillArea(ByRef Pic As PictureBox, ByVal NewColor As Long)

    Rem Create a solid brush
    mBrush = CreateSolidBrush(NewColor)
    Rem Select the brush into the hDC (device context)
    SelectObject Pic.hdc, mBrush
    Rem API uses pixels!
    Pic.ScaleMode = vbPixels
    Pic.AutoRedraw = True
    ExtFloodFill Pic.hdc, 0, 0, GetPixel(Pic.hdc, 0, 0), FLOODFILLSURFACE

    Rem Delete our new brush
    DeleteObject mBrush

End Function

Public Function GetButtonFaceColor() As Long
  GetButtonFaceColor = GetSysColor(COLOR_BTNFACE)
End Function
