VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGenericLeaf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database

Public Property Get IsRed() As Boolean: End Property
Public Property Let IsRed(ByVal Value As Boolean): End Property

Public Property Get Left() As IGenericLeaf: Set Left = myClass.Left: End Property
Public Property Set Left(ByVal Node As IGenericLeaf): Set myClass.Left = Node: End Property

Public Property Get Right() As IGenericLeaf: Set Right = myClass.Right: End Property
Public Property Set Right(ByVal Node As IGenericLeaf): Set myClass.Right = Node: End Property

Public Property Get Key() As IGeneric: End Property
Public Property Set Key(ByVal k As IGeneric): End Property

Public Property Get Value() As IGeneric: End Property
Public Property Set Value(ByVal v As IGeneric): End Property
Public Function AsPair() As GenericPair: End Function
