VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FieldListTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

'AccUnit:TestClass
'VBA-Inspector:Ignore

' AccUnit infrastructure for advanced AccUnit features. Do not remove these lines.
Implements SimplyVBUnit.ITestFixture
Implements AccUnit_Integration.ITestManagerBridge
Private TestManager As AccUnit_Integration.TestManager
Private Sub ITestManagerBridge_InitTestManager(ByVal NewTestManager As AccUnit_Integration.ITestManagerComInterface): Set TestManager = NewTestManager: End Sub
Private Function ITestManagerBridge_GetTestManager() As AccUnit_Integration.ITestManagerComInterface: Set ITestManagerBridge_GetTestManager = TestManager: End Function
Private Sub ITestFixture_AddTestCases(ByVal Tests As SimplyVBUnit.TestCaseCollector): TestManager.AddTestCases Tests: End Sub

'--------------------------------------------------------------------
' Tests
'--------------------------------------------------------------------

Public Sub ForEach()
   
   Dim f As Object
   Dim Fields As ACLibSqlTools.FieldList
   Dim i As Long

   Set Fields = SqlTools.FieldGenerator.FromNames("F1", "F2", "F3")

   For Each f In Fields.Items
      Assert.AreEqual Fields(i), f
      i = i + 1
   Next

End Sub

Public Sub FromStringArray()

   Dim Fields(2) As String
   Dim i As Long

   For i = LBound(Fields) To UBound(Fields)
      Fields(i) = "F" & i + 1
   Next
   
   Dim FieldList As ACLibSqlTools.FieldList
   Set FieldList = SqlTools.FieldGenerator.FromArray(Fields)

   Dim fld As ACLibSqlTools.Field
   i = 0
   For Each fld In FieldList.Items
      Assert.AreEqual Fields(i), fld.Name
      i = i + 1
   Next

End Sub

Public Sub FromStringParamArray()

   Dim Fields(2) As String
   Dim i As Long

   For i = LBound(Fields) To UBound(Fields)
      Fields(i) = "F" & i + 1
   Next
   
   Dim FieldList As ACLibSqlTools.FieldList
   Set FieldList = SqlTools.FieldGenerator.FromArray(ParamArrayToArray("F1", "F2", "F3"))

   Dim fld As ACLibSqlTools.Field
   i = 0
   For Each fld In FieldList.Items
      Assert.AreEqual Fields(i), fld.Name
      i = i + 1
   Next

End Sub

Private Function ParamArrayToArray(ParamArray f() As Variant) As Variant()
   ParamArrayToArray = f
End Function

'AccUnit:Ignore("Warum funktioniert das nicht?")
Public Sub FromFieldArray()

   Dim Fields(2) As ACLibSqlTools.Field
   Dim i As Long

   For i = LBound(Fields) To UBound(Fields)
      Set Fields(i) = SqlTools.FieldGenerator.Field("F" & i)
   Next

   Dim FieldList As ACLibSqlTools.FieldList
   Set FieldList = SqlTools.FieldGenerator.FromArray(Fields)

   Dim fld As ACLibSqlTools.Field
   i = 0
   For Each fld In FieldList.Items
      Assert.AreEqual Fields(i), fld.Name
      i = i + 1
   Next

End Sub

