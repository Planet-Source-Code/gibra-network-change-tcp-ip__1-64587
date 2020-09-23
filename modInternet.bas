Attribute VB_Name = "modInternet"
'---------------------------------------------------------------------------------------
' Module      : modInternet
' DateTime    : 24/02/2005 10.08
' Author      : Giorgio Brausi (vbcorner@vbcorner.net - http://www.vbcorner.net)
' Project     : none
' Purpose     : Two simple routines for Internet:
'             : SendEmail: send e-mail with default mailer (no attach)
'             : OpenUrl  : Open the URL with default browser
'
'---------------------------------------------------------------------------------------
Option Explicit

Private Declare Function ShellExecute _
    Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hWnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

  Const SW_SHOWNORMAL As Long = 1
  
  Rem Use this to add a CR in your body text
  Public Const gsCR = "%0d%0a"



'---------------------------------------------------------------------------------------
' Procedure   : SendEMail
' DateTime    : 24/02/2005 11.02
' Author      : Giorgio Brausi (vbcorner@vbcorner.net - http://www.vbcorner.net)
' Purpose     : Invia un email
' Descritpion :
' Comments    : Work with ANY mailer
'---------------------------------------------------------------------------------------
Public Sub SendEMail(ByVal hWnd As Long, ByVal sEmailAddress As String, Optional ByVal sSubject As String = "", Optional ByVal sBody As String = "")

  Const SW_SHOWNORMAL = 1
  Const SW_SHOW As Long = 5
  Dim sText As String
  Dim sAddedText As String
  Dim txtMainAddresses As String, txtCc As String, txtBcc As String
  Dim txtSubject As String, txtBody As String
  
  txtMainAddresses = sEmailAddress
  txtSubject = sSubject
  txtBody = sBody
  
  If Len(txtMainAddresses) Then
      sText = txtMainAddresses
  End If
  If Len(txtCc) Then
      sAddedText = sAddedText & "&CC=" & txtCc
  End If
  If Len(txtBcc) Then
      sAddedText = sAddedText & "&BCC=" & txtBcc
  End If
  If Len(txtSubject) Then
      sAddedText = sAddedText & "?Subject=" & txtSubject
  End If
  If Len(txtBody) Then
      sAddedText = sAddedText & "&Body=" & txtBody
  End If
  
  sText = "mailto:" & sText
  ' clean the added elements
  If Len(sAddedText) <> 0 Then
      ' there are added elements, replace the first
      ' ampersand with the question character
      Mid$(sAddedText, 1, 1) = "?"
  End If
  
  sText = sText & sAddedText
  
  If Len(sText) Then
    Call ShellExecute(hWnd, "open", sText, vbNullString, vbNullString, SW_SHOW)
  End If

End Sub

'---------------------------------------------------------------------------------------
' Procedure   : OpenUrl
' DateTime    : 24/02/2005 11.01
' Author      : Giorgio Brausi (vbcorner@vbcorner.net - http://www.vbcorner.net)
' Purpose     : Open web page <strURL>
' Descritpion : Use default browser
' Comments    : Work with any browser
'---------------------------------------------------------------------------------------
Public Sub OpenUrl(ByVal strURL As String)

  Call ShellExecute(0, "Open", strURL, 0&, 0&, SW_SHOWNORMAL)

End Sub






