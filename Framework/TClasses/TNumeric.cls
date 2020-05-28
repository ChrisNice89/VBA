VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TNumeric"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Enum BuildState
    Ready = 0
    [_New]
End Enum

Private Type Member
    Value As Currency
    Build As BuildState
End Type
Private myclass As Member

Public Function Build(ByVal Value As Currency) As TNumeric
    
    If Me Is TNumeric Then
        With myclass
            .Build = BuildState.[_New]
            .Value = Value
            Set n = New TNumeric
            .Value = 0
            .Build = Ready
        End With
    Else
        Call Err.Raise(5, TypeName$(Me), "Invalid instance use " & TypeName$(Me) & ".Build() to instatiate new " & TypeName$(Me) & " objects.")
    End If
    
End Function

Private Sub Class_Initialize()
    If Not (Me Is TNumeric) Then
        If myclass.Build = [_New] Then
            'myclass.Value =
        Else
            Call Err.Raise(5, TypeName$(Me), "Invalid instance use " & TypeName$(Me) & ".Build() to instatiate new " & TypeName$(Me) & " objects.")
        End If
    
    End If

End Sub
