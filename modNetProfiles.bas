Attribute VB_Name = "modNetProfiles"
'---------------------------------------------------------------------------------------
' Module      : modNetProfiles
' DateTime    : 13/06/2005 00.00
' Author      : Giorgio Brausi (vbcorner@vbcorner.net - http://www.vbcorner.net)
' Project     : Network_Change
' Purpose     :
' Descritpion :
' Comments    :
'---------------------------------------------------------------------------------------
Option Explicit
Public Declare Sub InitCommonControls Lib "comctl32.dll" ()
Public CIni As clsIni

Public Const INI_SECTION_PROFILES = "Profiles"

Public Enum enumPROFILE_MODE
  PROF_MODE_SAVE = 0
  PROF_MODE_EDIT = 1
  PROF_MODE_CLONE = 2
End Enum
#If False Then
  Public PROF_MODE_SAVE
  Public PROF_MODE_EDIT
  Public PROF_MODE_CLONE
#End If
Public eProfileMode As enumPROFILE_MODE

Public bOnFocus As Boolean
'---------------------------------------------------------------------------------------
' Procedure   : ProfileSave
' DateTime    : 16/06/2005 14.35
' Author      : Giorgio Brausi (vbcorner@vbcorner.net - http://www.vbcorner.net)
' Purpose     : Save a new or modified profile
' Descritpion :
' Comments    :
'---------------------------------------------------------------------------------------
Public Sub ProfileSave(ByVal sINIFileName As String, ByVal sProfileName, ByRef txtIP As Variant)
  Dim sMsg As String
  Dim mINI As clsIni, i As Integer, sTmp As String
  Set mINI = New clsIni
  
  mINI.FileName = sINIFileName
  
  If eProfileMode = PROF_MODE_SAVE Then
    Rem It's a new Profile: then verify if already exists
    Rem another profile with same name.
    Rem If Exists, ask to overwrite it
    If mINI.ReadString(INI_SECTION_PROFILES, sProfileName, "") <> "" Then
      sMsg = MLSGetString("0023")
      sMsg = Replace(sMsg, "%s", sProfileName) & vbCrLf & MLSGetString("0024")
      If MsgBox(sMsg, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    End If
  End If
  
  For i = 0 To 19
    sTmp = sTmp & "." & txtIP(i)
  Next i
  sTmp = Mid(sTmp, 2)
  mINI.WriteString INI_SECTION_PROFILES, sProfileName, sTmp
  
  Set mINI = Nothing
End Sub


'---------------------------------------------------------------------------------------
' Procedure   : ProfilesLoad
' DateTime    : 12/06/2005 12.59
' Author      : Giorgio Brausi (vbcorner@vbcorner.net - http://www.vbcorner.net)
' Purpose     : Read all Profiles from INI file and put them in ListBox
'             :
'---------------------------------------------------------------------------------------
Public Sub ProfilesLoad(ByVal sINIFileName As String, ByRef lst As ListBox)

  Dim mINI As clsIni
  Set mINI = New clsIni
  Dim sKey() As String, iCount As Long, i As Integer
  On Error GoTo ProfilesLoad_Error

  With mINI
    .FileName = sINIFileName
    .ReadSection .FileName, INI_SECTION_PROFILES, sKey(), iCount
    
    lst.Clear
    If iCount > 0 Then
      For i = 1 To iCount
        lst.AddItem sKey(i)
      Next i
    End If
  End With
  On Error GoTo 0
  Set mINI = Nothing
  Exit Sub

ProfilesLoad_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ")" & vbCrLf & "in procedure ProfilesLoad of Modulo modNetProfiles"
    Resume
End Sub
