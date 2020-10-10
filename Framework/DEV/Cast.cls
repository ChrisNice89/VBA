VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("Base")


Option Explicit
'@PredeclaredId

Public Function ToString(ByVal Instance As IGeneric) As GString: Set ToString = Instance: End Function
Public Function ToNumeric(ByVal Instance As IGeneric) As GNumeric: Set ToNumeric = Instance: End Function


Public Function SqlValue(ByVal Instance As IGeneric) As ISqlValue: Set SqlValue = Instance: End Function

Public Function ToArray(ByVal Instance As IGeneric) As GenericArray: Set ToArray = Instance: End Function
Public Function ToHashMap(ByVal Instance As IGeneric) As GenericMap: Set ToHashMap = Instance: End Function
Public Function ToList(ByVal Instance As IGeneric) As GenericList: Set ToList = Instance: End Function
Public Function ToSet(ByVal Instance As IGeneric) As GenericSortedSet: Set ToSet = Instance: End Function

Public Function ToCollection(ByVal Instance As IGeneric) As IGenericCollection: Set ToCollection = Instance: End Function
Public Function ToDictionary(ByVal Instance As IGeneric) As IGenericDictionary: Set ToDictionary = Instance: End Function
Public Function ToPair(ByVal Instance As IGeneric) As GenericPair: Set ToPair = Instance: End Function

