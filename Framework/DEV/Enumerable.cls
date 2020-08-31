VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Enumerable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("<T>Enumerable")
Option Compare Database

Option Explicit
'@PredeclaredId

Public Event Aggregate(ByRef Result As IGeneric, ByVal Current As IGeneric)

Private Type Member
    Sequence As IGenericEnumerator
End Type
Private myClass As Member
'Private Formats() As String

Private Type Internal
    Size As Long
    Pointer As LongPtr
End Type
Private Virtual As Internal

Public Property Get Sequence() As IGenericEnumerator
    Set Sequence = myClass.Sequence
End Property

Public Function Build(ByVal Sequence As IGenericEnumerator) As Enumerable
    
    If Sequence Is Nothing Then _
        Call Skynet.Throw(Me, "Public Function Build").NullPointer("Sequence is nothing")
      
    Set myClass.Sequence = Sequence
    Set Build = New Enumerable
    
End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myClass)
    Set myClass.Sequence = Enumerable.Sequence
End Sub

Public Sub Aggregate(ByRef Source As IGeneric, Optional ByVal IgnoreNull As Boolean = True)
    
    With myClass.Sequence
        Call .Reset
        If IgnoreNull Then
            Do While .MoveNext
                If Not .Current Is Nothing Then _
                    RaiseEvent Aggregate(Source, .Current)
            Loop
        Else
            Do While .MoveNext
                RaiseEvent Aggregate(Source, .Current)
            Loop
        End If
    End With
    
End Sub

