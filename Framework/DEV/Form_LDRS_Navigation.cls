VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_LDRS_Navigation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Private Type FormProperties
    AllowQuit As Boolean
End Type

Private mThisForm As FormProperties

Public Property Get ImageList() As MSComctlLib.ImageList
    'Set ImageList = Me.axImageList.Object
End Property

'Beim Start werden alle Abfragen Formulare im Hintergrund geöffnet
Private Sub Form_Load()
 
  
    Call Me.LoadSubform(Form_LDRS_OVERVIEW)
End Sub

'Beim Schließen werden alle Abfragen Formulare im Hintergrund geschlossen
Private Sub Form_Unload(Cancel As Integer)
  
    DoCmd.ShowToolbar "Ribbon", acToolbarYes
    If mThisForm.AllowQuit Then
        Application.Quit
    End If
End Sub

Private Sub Navi_Exit_Click()

    Dim strDialogue As String
    
    Select Case True
        Case 1 'Deutsch
            strDialogue = "Soll das LDRS geschlossen werden?"
        Case 2 'English
            strDialogue = "Do you want to leave LDRS?"
    End Select
            
    If MsgBox(strDialogue, vbSystemModal + vbQuestion + vbYesNo, "iCAT©    Exit") = vbYes Then
        mThisForm.AllowQuit = True
        DoCmd.Close acForm, Me.Name
    End If
    
End Sub

Private Sub Navi_Overview_Click()
    Call Me.LoadSubform(Form_LDRS_OVERVIEW)
End Sub

Public Sub LoadSubform(ByVal Form As Access.Form)

    Dim ctlSubform As Access.SubForm
    Set ctlSubform = Me.Controls(Me.Navigation.SubForm)
    
    Call SetLanguage
  
    If Form.Name = "Overview" Then
        ctlSubform.SourceObject = "Overview"
  
    Else
        ctlSubform.SourceObject = Form.Name
    End If
   
End Sub

Private Sub SetLanguage()

    Select Case True
        Case 1 'Deutsch
            With Me.Navigation.Tabs
                .Item("Navi_Overview").Caption = "Übersicht"
                .Item("Navi_Assessment").Caption = "Analysen"
            End With
        
        Case 2 'English
            With Me.Navigation.Tabs
                .Item("Navi_Overview").Caption = "Overview"
                .Item("Navi_Assessment").Caption = "Analysis"
            End With
    End Select
    
End Sub


