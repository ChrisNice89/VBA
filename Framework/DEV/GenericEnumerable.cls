VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericEnumerable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("<T>Enumerable")
Option Compare Database

Option Explicit
'@PredeclaredId

Public Event Aggregate(ByRef Result As IGeneric, ByVal Current As IGeneric, ByVal NullsIgnored As Boolean)
Public Event Contains(ByRef Contains As Boolean, ByVal Current As IGeneric)


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

Public Function Build(ByVal Sequence As IGenericEnumerator) As GenericEnumerable
    
    If Sequence Is Nothing Then _
        Call Skynet.Throw(Me, "Public Function Build").NullPointer("Sequence is nothing")
      
    Set myClass.Sequence = Sequence
    Set Build = New GenericEnumerable
    
End Function

Private Sub Class_Initialize()
    Virtual.Pointer = VarPtr(myClass)
    Set myClass.Sequence = GenericEnumerable.Sequence
End Sub

Private Sub Check(ByVal Source As IPrototype, ByVal Caller As String)
    
    If (Source Is Nothing) Then _
        Set Source = Source.Default 'Call Skynet.Throw(Me, Caller).NullPointer("Key is nothing")

    If myClass.Sequence Is Nothing Then _
        Call Skynet.Throw(Me, Caller).NullPointer("Sequence is nothing")
      
End Sub

Public Function Aggregate(ByVal Result As IPrototype, Optional ByVal IgnoreNull As Boolean = True) As IGeneric
    
    Call Check(Result, "Public Sub Aggregate")
    
    With myClass.Sequence
        Call .Reset
        If IgnoreNull Then
            Do While .MoveNext
                If Not .Current Is Nothing Then _
                    RaiseEvent Aggregate(Result, .Current, IgnoreNull)
            Loop
        Else
            Do While .MoveNext
                RaiseEvent Aggregate(Result, .Current, IgnoreNull)
            Loop
        End If
    End With
    
    Set Aggregate = Result
    
End Function

Public Function Contains(ByVal Item As IGeneric) As Boolean
   
    With myClass.Sequence
        Call .Reset
        If Item Is Nothing Then
            Do While .MoveNext
                If .Current Is Nothing Then _
                   Contains = True: Exit Function
            Loop
        Else
            Do While .MoveNext And Not Contains
                If Item.Equals(.Current) Then _
                   Contains = True: Exit Function
            Loop
        End If
    End With
    
End Function
