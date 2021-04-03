VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericDelegate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

Implements IGenericComparer

Public Event Action(ByRef Element As IGeneric)
Public Event Predicate(ByRef Result As Boolean, ByVal Element As IGeneric)
Public Event Compare(ByRef Result As CompareResult, ByVal A As IGeneric, ByVal B As IGeneric)
Public Event Equals(ByRef Result As Boolean, ByVal A As IGeneric, ByVal B As IGeneric)

Public Function Of(ByRef Delegate As GenericDelegate)
Attribute Of.VB_UserMemId = 0
    Set Delegate = New GenericDelegate
    Set Of = Delegate
End Function

Private Function IGenericComparer_Compare(ByVal A As IGeneric, ByVal B As IGeneric) As CompareResult

    If A Is B Then _
        IGenericComparer_Compare = CompareResult.IsEqual: Exit Function
        
    If A Is Nothing Then _
        IGenericComparer_Compare = CompareResult.islower: Exit Function
    
    If B Is Nothing Then _
        IGenericComparer_Compare = CompareResult.isgreater: Exit Function
    
    RaiseEvent Compare(IGenericComparer_Compare, A, B)

End Function

Private Function IGenericComparer_Equals(ByVal A As IGeneric, ByVal B As IGeneric) As Boolean

    If A Is B Then _
        IGenericComparer_Equals = True: Exit Function
        
    If A Is Nothing Or B Is Nothing Then _
        Exit Function
    
    RaiseEvent Equals(IGenericComparer_Equals, A, B)

End Function

Private Function IGenericComparer_GetHashCode(ByVal Obj As IGeneric) As Long



End Function
