Attribute VB_Name = "modLanguageSupport"
Option Explicit

' -----------------------------------------------------------------
' This module was made by Multi-Language Support Add-in for VB,
' by Giorgio Brausi (gibra)
' Contact me by e-mail: vbcorner@vbcorner.net
' Web site: http://www.vbcorner.net
' -----------------------------------------------------------------
' Your languages there is not? Want to add it?
' To add new language files:
' 1 - Copy and rename ENGLISH.LNG to your language (i.e. FRENCH.LNG)
' 2 - Tranlsate strings right to '=' character
'     IMPORTANT: leave unchanged the '%s' sequence!
' 3 - Send me the new language, so i can update the project
'
' Note: You have not to modify the project to use your language!
'       Simply put it together other LNG files, NC will recognize
'       it automatically and will add it to the menu. ;-))
' -----------------------------------------------------------------
Public gsLanguageFile As String  '/language file (i.e. english.lng, italian.lng,...)

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

'/ Update all controls properties with the current language
Public Sub MLSLoadLanguage(Form As Form)
    Dim Obj As Object
    Dim sFileName As String, a As String
    
    On Error Resume Next
    
    If Right(App.Path, 1) = "\" Then
        sFileName = App.Path & gsLanguageFile & ".lng"
    Else
        sFileName = App.Path & "\" & gsLanguageFile & ".lng"
    End If

    Rem Load Caption for Form, if there
    If Len(Form.Caption) > 0 Then
        Form.Caption = MLSReadINI(sFileName, CStr(Form.Name), CStr(Form.Name) & ".Caption")
    End If

    Form.ToolTipText = MLSReadINI(sFileName, CStr(Form.Name), CStr(Form.Name) & ".ToolTipText")
    Form.Tag = MLSReadINI(sFileName, CStr(Form.Name), CStr(Form.Name) & ".Tag")

    '/ Load properties for objects
    For Each Obj In Form
        Dim bHasIndex As Boolean '/ to check if has Index property
        a$ = ""
        '/ If is not a matrix return a error code 343
        bHasIndex = (Obj.Index >= 0)
        If Err.Number = 343 Then     '/ The object is not a matrix
            bHasIndex = False
            Err.Clear
        End If

        Rem Get Caption property
        If bHasIndex Then '/ This is a matrix
            a$ = MLSReadINI(sFileName, CStr(Form.Name), Obj.Name & "(" & Obj.Index & ").Caption")
        Else
            a$ = MLSReadINI(sFileName, CStr(Form.Name), Obj.Name & ".Caption")
        End If
        If a$ <> "" Then
            Obj.Caption = a$
        End If

        Rem Get ToolTipText property
        a$ = ""
        If bHasIndex Then '/ This is a matrix
            a$ = MLSReadINI(sFileName, CStr(Form.Name), Obj.Name & "(" & Obj.Index & ").ToolTipText")
        Else
            a$ = MLSReadINI(sFileName, CStr(Form.Name), Obj.Name & ".ToolTipText")
        End If
        If a$ <> "" Then
            Obj.ToolTipText = a$
        End If

        '/ Get Tag property
        a$ = ""
        If bHasIndex Then '/ This is a matrix
            a$ = MLSReadINI(sFileName, CStr(Form.Name), Obj.Name & "(" & Obj.Index & ").Tag")
        Else
            a$ = MLSReadINI(sFileName, CStr(Form.Name), Obj.Name & ".Tag")
        End If
        If a$ <> "" Then
            Obj.Tag = a$
        End If

        '/ check properties for SSTab control: this control has
        '/ Caption for each tab, named TabCaption
        If Obj.Tabs Then
            If Err = 0 Then
                a$ = ""
                Dim nT As Integer
                '/ find the caption for each Tab
                For nT = 0 To Obj.Tabs
                     Obj.TabCaption(nT) = MLSReadINI(sFileName, CStr(Form.Name), Obj.Name & ".TabCaption(" & nT & ")")
                Next nT
            Else
                Err.Clear
            End If
        End If


        DoEvents

    Next
End Sub
'/ Load a translate string from [Strings] section of current language
Public Function MLSGetString(KeyName As String) As String
Dim sFileName As String

    On Error Resume Next
    
    Rem Get the language filename on 'real-time'
    If gsLanguageFile = "" Then
        gsLanguageFile = MLSReadINI(App.Path & "\" & "LangSetting.ini", "Language", "CurrentLanguage")
    End If
    If Right(App.Path, 1) = "\" Then
        sFileName = App.Path & gsLanguageFile & ".lng"
    Else
        sFileName = App.Path & "\" & gsLanguageFile & ".lng"
    End If
    MLSGetString = MLSReadINI(sFileName, "Strings", KeyName$)
End Function
Public Function MLSReadINI(File$, SectionName$, KeyName$) As String
Dim Value As String * 1024, i As Long

    '/ If the file INI is large than 64k, GetPrivateProfileString fail
    '/ then I use Open method to retrieve the string:
    If FileLen(File$) > 64000 Then

        '/ Use Open
        Dim numFile As Integer, sThisLine As String, sTmp As String
        numFile = FreeFile
        Open File$ For Input As #numFile
        Do While Not EOF(numFile)
            Line Input #numFile, sThisLine
            If Left(sThisLine, Len(KeyName$)) = KeyName$ Then
                sTmp = Mid(sThisLine, Len(KeyName$) + 2)
                MLSReadINI = Mid(sTmp, 2, Len(sTmp) - 2)
                Close #numFile
                Exit Function
            End If
        Loop
        Close #numFile

    Else

        i = GetPrivateProfileString(SectionName$, KeyName$, "", Value, 512, File$)
        MLSReadINI = Left$(Value, InStr(Value, Chr$(0)) - 1)

    End If

End Function
Public Function MLSWriteINI(File$, SectionName$, KeyName$, NewValue$) As Long
    MLSWriteINI = WritePrivateProfileString(SectionName$, KeyName$, NewValue$, File$)
End Function

'---------------------------------------------------------------------------------------
' Procedure   : MLSGetProperty
' DateTime    : 19/07/2005 10.42
' Author      : Giorgio Brausi (vbcorner@vbcorner.net - http://www.vbcorner.net)
' Purpose     : Get a string from a specific property
' Descritpion : sFormName: the Name property of the form (i.e. frmMain)
'               Tips: You can set this parameter with:  Me.Name
'               sPropertyName: Name property to search for (i.e. frmMain.Caption)
' Comments    :
'---------------------------------------------------------------------------------------
Public Function MLSGetProperty(ByVal sFormName As String, ByVal sPropertyName As String) As String
Dim sFileName As String

    On Error Resume Next
    
    Rem Get the language filename on 'real-time'
    If gsLanguageFile = "" Then
        gsLanguageFile = MLSReadINI(App.Path & "\" & "LangSetting.ini", "Language", "CurrentLanguage")
    End If
    If Right(App.Path, 1) = "\" Then
        sFileName = App.Path & gsLanguageFile & ".lng"
    Else
        sFileName = App.Path & "\" & gsLanguageFile & ".lng"
    End If
    
   
    MLSGetProperty = MLSReadINI(sFileName, sFormName, sPropertyName)
  
End Function
