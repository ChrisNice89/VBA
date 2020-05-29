VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPrototype"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Property Get VirtualPointer() As LongPtr: End Property
Public Property Get Size() As Long: End Property
Public Property Get NewInstance() As IPrototype: End Property

Private Sub Class_Initialize()
    Call Skynet.Push(Me, "Private Sub Class_Initialize()").InterfaceInstantiation.Raise
End Sub
