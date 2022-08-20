VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGenericPredicate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder "Base.Abstract"
'@Exposed
'@PredeclaredId
'@Interface
Option Explicit

'@DefaultMember
Public Function Match(ByVal Element As Object) As Boolean: End Function
Attribute Match.VB_UserMemId = 0
