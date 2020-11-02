VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Table"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("Database")
Option Compare Database

Option Explicit
'@PredeclaredId

Public Enum PortfolioFields
    [_First]
'    Nummer
'    Name
'    Prüfer
'    Tranche
'    Prüfungsschwerpunkt
'    Auswahlgrund
'    Datum
'    Kundennummer
'    Kundenname
'    Ratingverfahren
'    Ratingnote
'    Ratingdatum
'    Risikovolumen
'    Inanspruchnahme
'    Blankovolumen
'    EWB
'    Kundennummer
'    Kontonummer
'    Produktgruppe
'    Produkttyp
'    Sollzins
'    Limit (extern)
'    Limit (intern)
'    Inanspruchnahme
'    Überziehungsdauer
'    Geber -Nummer
'    Geber -Name
'    Nummer
'    Sicherheitenart
'    Immobilien -Nummer
'    Objektart
'    BLW -Auslauf
'    Anrechnung
'    verfügbar
    [_Last]
End Enum

Public Enum NormalFields
    [_First]

    [_Last]
End Enum


Private Type Member
    Portfolio() As String
    Normal() As String
End Type
Private myclass As Member

Private Sub Class_Initialize()
    With myclass
        ReDim Portfolio(PortfolioFields.[_First] To PortfolioFields.[_Last])
    
    End With
End Sub

Public Property Get Portfolio(ByVal Field As PortfolioFields) As String
    Portfolio = myclass.Portfolio(Field)
End Property
