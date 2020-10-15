VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database

Implements IGeneric

Private Type Member
   Parameter As GenericOrderedMap
   Values  As GenericOrderedMap
End Type
Private myclass As Member

Public Property Get ParameterByIndex(ByVal Index As Long) As IGeneric
    Set ParameterByIndex = myclass.Parameter.ElementAt(Index)
End Property

Public Property Get Parameter(ByVal Name As GString) As IGeneric
    
    Dim P As IGeneric
    
    If myclass.Parameter.TryGetValue(Name, P) Then
        Set Parameter = P
    Else
        'Error
    End If
    
End Property

Public Property Get GetValues(ByVal Name As GString) As IGenericCollection
    
    Dim Values As GenericSortedSet
    
    If myclass.Values.TryGetValue(Name, Values) Then
        Set GetValues = Values
    Else
        'Error
    End If
End Property

Public Property Get GetValuesByIndex(ByVal Index As Long) As GenericSortedSet
    
    Dim Values As GenericSortedSet
    
    Set Values = myclass.Values.ElementAt(Index)
    
    If Not Values Is noting Then
        Set GetValuesByIndex = Tree
    Else
        'Error
    End If
End Property

Public Function CreateParameter( _
                                    ByVal Name As String, _
                                    ByVal Value As ISqlValue) As Object

    Dim P As ADODB.Parameter
    Set P = New ADODB.Parameter
    
    With P
        .Direction = adParamInput
        .Name = Name
        .Attributes = adParamNullable
        
        If Value.IsNumeric Then
            .Precision = 4
            .NumericScale = 4
        Else
            .Size = Len(Value.Primitive)
        End If
        .Value = Value.Primitive
    End With
    
    Set CreateParameter = P

End Function

Public Sub Initialize(ByVal Number As Long)
    Set myclass.Parameter = GenericOrderedMap.Build(Number)
    Set myclass.Values = GenericOrderedMap.Build(Number)
End Sub

Public Function AddWithValues( _
                                ByVal Name As GString, _
                                ParamArray Values() As Variant)

    Dim Value As IGeneric
    
    If IsMissing(Values) Then _
        Call Skynet.Throw(Me, "Public Function AddWithValues_2").NullPointer("ParamArray GenericValues are missing")
    
    Dim SortedSet As GenericSortedSet
    Set SortedSet = GenericSortedSet.Build
    
    For i = LBound(Values) To UBound(Values)
        If Not TypeOf Values(i) Is ISqlValue Then
            'error
        End If
        Call SortedSet.Add(Values(i))
    Next
        
    With myclass
        Call .Values.Add(Name, SortedSet)
        Call .Parameter.Add(Name, Values(1))
    End With
    
End Function

'IGeneric

Private Property Get IGeneric_IsMutable() As Boolean

End Property

Private Function IGeneric_IsRelatedTo(ByVal Other As IGeneric) As Boolean

End Function

Private Function IGeneric_ToString() As String

End Function

Private Function IGeneric_Clone() As IGeneric

End Function

Private Function IGeneric_Equals(ByVal Other As IGeneric) As Boolean

End Function
 
Private Function IGeneric_CompareTo(ByVal Other As IGeneric) As ComparerEnum

End Function

Private Sub IGeneric_Dispose()

End Sub

Private Property Get IGeneric_HashValue() As Long

End Property
