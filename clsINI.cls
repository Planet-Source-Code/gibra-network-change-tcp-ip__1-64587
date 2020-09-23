VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
  (ByVal lpApplicationName As String, _
  ByVal lpKeyName As Any, _
  ByVal lpString As Any, _
  ByVal lpFileName As String) As Long

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
  (ByVal lpApplicationName As String, _
  ByVal lpKeyName As Any, _
  ByVal lpDefault As String, _
  ByVal lpReturnedString As String, _
  ByVal nSize As Long, _
  ByVal lpFileName As String) As Long

Private Const MAXLEN As Integer = 255
Private Const z As Integer = 0

Private pIniFileName As String
Private pSection As String
'---------------------------------------------------------------------------------------
' Procedure   : ReadSection
' DateTime    : 01/09/2004 12.09
' Author      : Giorgio Brausi (vbcorner@vbcorner.net - http://www.vbcorner.net)
' Purpose     : Read the entire section
' Descritpion : Return a string array (sKey) and the number of entries (iCount)
' Comments    : IMPORTANT: before call the procedure, on calling form need to declare:
'
'               Dim sKey() As String, iCount As Long
'
'               This paramters will be passed ByRef!
'---------------------------------------------------------------------------------------
Public Function ReadSection(IniFile As String, sSectionName As String, ByRef sKey() As String, ByRef iCount As Long) As Boolean
Dim sSection As String
Dim iPos As Long
Dim iNextPos As Long
Dim sCur As String
    
    ' ================================================
    ' Get all section entries
    ' ================================================
    Dim sBuf As String
    Dim iSize As String
    Dim iRetCode As Integer
    Dim sDefault As String
    Dim sPath As String
    sPath = IniFile
    sBuf = Space$(8192)
    iSize = Len(sBuf)
    iRetCode = GetPrivateProfileString(sSectionName, 0&, sDefault, sBuf, iSize, sPath)
    If (iSize > 0) Then
        sSection = Left$(sBuf, iRetCode)
    Else
        sSection = ""
    End If
    
    iCount = 0
    Erase sKey
    If (Len(sSection) > 0) Then
        iPos = 1
        iNextPos = InStr(iPos, sSection, Chr$(0))
        Do While iNextPos <> 0
            sCur = Mid$(sSection, iPos, (iNextPos - iPos))
            If (sCur <> Chr$(0)) Then
                iCount = iCount + 1
                ReDim Preserve sKey(1 To iCount) As String
                sKey(iCount) = Mid$(sSection, iPos, (iNextPos - iPos))
                iPos = iNextPos + 1
                iNextPos = InStr(iPos, sSection, Chr$(0))
            End If
        Loop
    End If
    
    Rem iCount contains the number of entries
    Rem sKey() contains all the Name entries
    
End Function
Public Property Let FileName(ByVal Value As String)
    Dim FileNum%
    Const sDot = "."
    Const sBackslash = "\"
    Const sExtension = ".ini"
    
    pIniFileName = Value
    If InStr(pIniFileName, sDot) = z Then pIniFileName = pIniFileName & sExtension
    If InStr(pIniFileName, sBackslash) = z Then pIniFileName = App.Path & sBackslash & pIniFileName
  
    If Dir(pIniFileName) = "" Then
        Rem Create the file, if not exists
        FileNum = FreeFile
        Open pIniFileName For Append As FileNum
        Close #FileNum
    End If
    
    
    
End Property

Public Property Get FileName() As String
  FileName = pIniFileName
End Property

Public Property Get Section() As String
  Section = pSection
End Property

Public Property Let Section(ByVal Value As String)
  pSection = Value
End Property

Public Function WriteString(ByVal sSection As String, ByVal sKey As String, ByVal Value As String) As Long
    WriteString = WritePrivateProfileString(sSection, sKey, Value, pIniFileName)
End Function

'---------------------------------------------------------------------------------------
' Procedure   : ReadBool
' DateTime    : 21/02/2005 15.00
' Author      : Giorgio Brausi (vbcorner@vbcorner.net - http://www.vbcorner.net)
' Purpose     : Read a STRING value and return a BOOLEAN value
' Descritpion :
' Comments    :
'---------------------------------------------------------------------------------------
Public Function ReadBool(ByVal sSection As String, ByVal sKey As String, ByVal Default As Boolean) As Boolean
  Dim Value As String, lRet As Long, ReadString As String
  If sSection = "" Then Exit Function
  If sKey = "" Then Exit Function
  
  Value = String$(MAXLEN, z)
  lRet = GetPrivateProfileString(sSection, sKey, Default, Value, MAXLEN, pIniFileName)
  ReadString = Left(Value, lRet)
  
  Const sFalso = "Falso"
  Const sZero = "0"
  
  If ReadString = sFalso Or ReadString = sZero Or ReadString = "" Then
    ReadBool = False
  Else
    ReadBool = True
  End If
    
End Function




'---------------------------------------------------------------------------------------
' Procedure   : ReadInt
' DateTime    : 21/02/2005 15.01
' Author      : Giorgio Brausi (vbcorner@vbcorner.net - http://www.vbcorner.net)
' Purpose     : Read a STRING value and return a INTEGER value
' Descritpion :
' Comments    :
'---------------------------------------------------------------------------------------
Public Function ReadInt(ByVal sSection As String, ByVal sKey As String, ByVal Default As Integer) As Integer
  Dim Value As String, lRet As Long
  If sSection = "" Then Exit Function
  If sKey = "" Then Exit Function
  
  Value = String$(MAXLEN, z)
  lRet = GetPrivateProfileString(sSection, sKey, Default, Value, MAXLEN, pIniFileName)
  
  If lRet = 0 Then
        ReadInt = Default
  Else
        ReadInt = Val(Left(Value, lRet))
  End If
  
End Function

Public Function ReadString(ByVal sSection As String, ByVal sKey As String, ByVal Default As String) As String
  Dim Value As String, lRet As Long
  If sSection = "" Then Exit Function
  If sKey = "" Then Exit Function
  
  Value = String$(MAXLEN, z)
  lRet = GetPrivateProfileString(sSection, sKey, Default, Value, MAXLEN, pIniFileName)
  
  If lRet = 0 Then
        ReadString = Default
  Else
        ReadString = Left(Value, lRet)
  End If
    
End Function

Rem Delete entry from INI file
Public Function DeleteKey(ByVal sSection As String, ByVal sKey As String) As Long
    Dim sIniFile As String
    sIniFile = FileName
    DeleteKey = WritePrivateProfileString(sSection, sKey, 0&, sIniFile)
    
End Function

Rem Delete entire section from INI file
Public Function DeleteSection(ByVal sSection As String) As Long
  Dim sFileIni As String
  sFileIni = FileName
  DeleteSection = WritePrivateProfileString(sSection, 0&, 0&, sFileIni)

End Function

