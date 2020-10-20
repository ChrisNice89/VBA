VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_LDRS_OVERVIEW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Type FormProperties
    Form As Form_LDRS_OVERVIEW
End Type
Private mThisForm As FormProperties
'

Private Sub Form_Load()
    Set mThisForm.Form = Me
   
   
End Sub

Private Function Blank() As FormProperties: End Function
Private Sub Form_Unload(Cancel As Integer)
    mThisForm = Blank

End Sub

Public Sub Show(ByVal GeberNr As Double, ByVal SiNr As Integer, ByVal ImmoNr As Integer)
'Übergabeparameter: GeberNr: Feld 'Geber Nr' aus Abfrage 'qGO_OCollateral'
'Übergabeparameter: SiNr: Feld 'Nummer' aus Abfrage 'qGO_OCollateral'
'Übergabeparameter: ImmoNr: Feld 'Immo-Nr' aus Abfrage 'qGO_OCollateral'
'                           Bei Implementierung beachten, dass ImmoNr NULL sein kann
    Screen.MousePointer = 11
    With mThisForm.Form
        .Modal = True
        .Visible = False
    End With
    
   
    Screen.MousePointer = 0
    mThisForm.Form.Visible = True
    
End Sub

