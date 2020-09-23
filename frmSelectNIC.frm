VERSION 5.00
Begin VB.Form frmSelectNIC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selezione Scheda di Rete"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6210
   Icon            =   "frmSelectNIC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkAllNIC 
      Caption         =   "Attiva su Tutte le Schede di Rete"
      Height          =   255
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   4875
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Conferma"
      Default         =   -1  'True
      Height          =   345
      Left            =   3420
      TabIndex        =   2
      Top             =   3420
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Chiudi"
      Height          =   345
      Left            =   4800
      TabIndex        =   3
      Top             =   3420
      Width           =   1245
   End
   Begin VB.ListBox lstNIC 
      Height          =   2400
      Left            =   120
      TabIndex        =   1
      Top             =   900
      Width           =   5955
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   0
      TabIndex        =   5
      Top             =   3900
      Width           =   6195
   End
   Begin VB.Label lblNIC 
      AutoSize        =   -1  'True
      Caption         =   "Seleziona la Sceda di Rete"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   660
      Width           =   1920
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5280
      Picture         =   "frmSelectNIC.frx":030A
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmSelectNIC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module      : frmSelectNIC
' DateTime    : 08/10/2005 14.45
' Author      : Roberto Doretto
' Project     : Network_Change
' Purpose     : Choose your Network Card
' Descritpion :
' Comments    : from v. 1.3.0
'---------------------------------------------------------------------------------------
Option Explicit

Dim pSelectedNIC As String

Private Sub chkAllNIC_Click()
    If chkAllNIC.Value = vbChecked Then
        lstNIC.Enabled = False
        pSelectedNIC = "*ALL"
    Else
        lstNIC.Enabled = True
        pSelectedNIC = ""
    End If
End Sub

Private Sub chkAllNIC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblStatus.Caption = MLSGetString("0039") ' MLS-> "Attiva su Tutte le Schede di Rete"
End Sub

Private Sub cmdCancel_Click()
  pSelectedNIC = ""
  Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim sMsg As String
    Rem If ALL cards has not been selected
    If chkAllNIC.Value = vbUnchecked Then
        If lstNIC.ListIndex = -1 Then
            Rem Warning: select the card, first
            sMsg = MLSGetString("0041") ' MLS-> "Selezionare prima una Scheda di Rete"
            MsgBox sMsg, vbCritical
            Exit Sub
        Else
            pSelectedNIC = lstNIC.Text
        End If
    End If
    Unload Me
End Sub

Private Sub cmdSave_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblStatus.Caption = MLSGetString("0037") ' MLS-> "Conferma Selezione"
End Sub

Private Sub Form_Activate()

    Rem If ONE card exists only, then :
    Rem  0. If the gbAutoSelectIfOnlyCardExists
    Rem  1. When form is open, select it automatically
    Rem  2.
    If Not gbAutoSelectIfOnlyCardExists Then Exit Sub
    
    If lstNIC.ListCount = 1 Then
      lstNIC.ListIndex = 0
      pSelectedNIC = lstNIC.Text
      Unload Me
    End If
    
End Sub

Private Sub Form_Load()
    Dim i As Integer
    MLSLoadLanguage Me '<- Add by Multi-Languages Support Add-in
  
    RetrieveLstNIC
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure   : RetrieveLstNIC
' DateTime    : 08/10/2005 15.34
' Author      : Roberto Doretto (vbcorner@vbcorner.net - http://www.vbcorner.net)
' Purpose     : Retrieve and list the network cards in your computer
'---------------------------------------------------------------------------------------

Private Sub RetrieveLstNIC()
    Dim strComputer As String
    Dim objWMIService As Object
    Dim colItems As Object
    Dim objItem As Object
    Dim NICDescription As String
    Dim pos As Integer

    On Error Resume Next
    strComputer = "."
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
    Set colItems = objWMIService.ExecQuery( _
        "SELECT * FROM Win32_NetworkAdapterConfiguration where IPEnabled=TRUE", , 48)
        
    Rem Add all network cards (NIC) found to ListBox
    For Each objItem In colItems
        pos = InStr(1, objItem.Caption, "] ")
        NICDescription = MLSGetString("0042") ' NIC not found
        If pos Then
            NICDescription = Mid$(objItem.Caption, pos + 2)
        End If
        lstNIC.AddItem NICDescription
    Next
    
    Rem Release objects
    Set objItem = Nothing
    Set colItems = Nothing
    Set objWMIService = Nothing
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblStatus.Caption = ""
End Sub

Private Sub lstNIC_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblStatus.Caption = MLSGetString("0038") ' MLS-> "Click: Seleziona la Scheda di Rete"
End Sub

Public Property Get SelectedNIC() As String
    SelectedNIC = pSelectedNIC
End Property

