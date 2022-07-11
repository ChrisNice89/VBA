VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericIndexIterator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule AssignedByValParameter
'@PredeclaredId
'@Folder "Collections"
Option Explicit

Implements IGenericIterator

Private Const NOT_SPECIFIED As Long = -1 'Dont change!

Private Type Member
    Source As IGenericCollection
    Index As Long
    Remaining As Long
    Current As Object
    Reverse As Boolean
End Type
Private This As Member

Public Property Get Index() As Long: Index = This.Index: End Property
Public Property Get Remaining() As Long: Remaining = This.Remaining: End Property
Public Property Get Reverse() As Boolean: Reverse = This.Reverse: End Property
Public Property Get Source() As IGenericCollection: Set Source = This.Source: End Property

Public Function Build( _
                        ByVal Collection As IGenericCollection, _
                        Optional ByVal BeginAt As Long = NOT_SPECIFIED, _
                        Optional ByVal Count As Long = NOT_SPECIFIED, _
                        Optional ByVal Reverse As Boolean = False) As IGenericIterator
    
    If (Collection Is Nothing) Then _
        Call GenericError.Build(Me, "Public Function Build").InvalidInput("(Count - BeginAt < Count)")
    
    If (BeginAt = NOT_SPECIFIED) Then
        BeginAt = 0
    ElseIf (BeginAt < 0) Then
        Call GenericError.Build(Me, "Public Function Build").ArgOutOfRange("BeginAt < 0")
    End If
    
    If (Count = NOT_SPECIFIED) Then
        Count = Collection.Count - BeginAt
    ElseIf (Count < 0) Then
        Call GenericError.Build(Me, "Public Function Build").ArgOutOfRange("Count < 0")
    End If
    
    If (Collection.Count - BeginAt < Count) Then _
        Call GenericError.Build(Me, "Public Function Build").InvalidInput("(Count - BeginAt < Count)")
    
    If Not Reverse Then
        This.Reverse = False
        This.Index = BeginAt
    Else
        This.Reverse = True
        This.Index = Count - 1
    End If
    
    This.Remaining = Count
    Set This.Source = Collection
    
    Set Build = New GenericIndexIterator
    This = EmptyMember
    
End Function

'@Ignore NonReturningFunction, EmptyMethod
Private Function EmptyMember() As Member: End Function

Private Sub Class_Initialize()
    
    If Me Is GenericIndexIterator Then _
        Exit Sub
        
    With GenericIndexIterator
        This.Index = .Index
        This.Remaining = .Remaining
        This.Reverse = .Reverse
        Set This.Source = .Source
    End With
    
End Sub

Private Property Get IGenericIterator_Current() As Object: Set IGenericIterator_Current = This.Current: End Property

Private Function IGenericIterator_HasNext(Optional Out As Object) As Boolean
    
    If (This.Remaining > 0) Then
        
        Set This.Current = This.Source.GetAt(This.Index)
        Set Out = This.Current
        
        If (This.Reverse = False) Then
            This.Index = This.Index + 1
        Else
            This.Index = This.Index - 1
        End If
        
        This.Remaining = This.Remaining - 1
        IGenericIterator_HasNext = True
        Exit Function
    
    End If
   
    If This.Source Is Nothing Then _
        Call GenericError.Build(Me, "Private Function IGenericIterator_HasNext").InvalidIterator

    Set This.Source = Nothing
    
End Function


