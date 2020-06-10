VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestTab1DataSourceTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

'AccUnit:TestClass
Private TestTab1 As TestTab1DataSource

' AccUnit infrastructure for advanced AccUnit features. Do not remove these lines.
Implements SimplyVBUnit.ITestFixture
Implements AccUnit_Integration.ITestManagerBridge
Private TestManager As AccUnit_Integration.TestManager
Private Sub ITestManagerBridge_InitTestManager(ByVal NewTestManager As AccUnit_Integration.ITestManagerComInterface): Set TestManager = NewTestManager: End Sub
Private Function ITestManagerBridge_GetTestManager() As AccUnit_Integration.ITestManagerComInterface: Set ITestManagerBridge_GetTestManager = TestManager: End Function
Private Sub ITestFixture_AddTestCases(ByVal Tests As SimplyVBUnit.TestCaseCollector): TestManager.AddTestCases Tests: End Sub

'--------------------------------------------------------------------
' Setup/Teardown
'--------------------------------------------------------------------

Public Sub Setup()
   Set TestTab1 = New TestTab1DataSource
End Sub

Public Sub Teardown()
   Set TestTab1 = Nothing
End Sub

'--------------------------------------------------------------------
' Tests
'--------------------------------------------------------------------

' AccUnit:Tags(Example, Simple Test)
Public Sub F1text_GetName()
   
   Const Expected As String = "F1text"
   Dim Actual As String
   
   Actual = TestTab1.F1text.Name
   ' Assert
   Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub F1text_FirstRecord_GetValue()
   
   Const Expected As String = "a"
   Dim Actual As Variant
   
   TestTab1.Source.LoadData
   Actual = TestTab1.F1text.Value

   Assert.That Actual, Iz.EqualTo(Expected)

End Sub

Public Sub F1text_MoveNext_GetValue()
   
   Const Expected As String = "b"
   Dim Actual As Variant
   
   TestTab1.Source.LoadData
   TestTab1.Source.MoveNext

   Actual = TestTab1.F1text.Value

   Assert.That Actual, Iz.EqualTo(Expected)

End Sub

Public Sub F1text_WithFilter_Id2()
   
   Const Expected As String = "b"
   Dim rst As ADODB.Recordset
   Dim Actual As Variant

   Dim Filter As ACLibSqlTools.ICondition
   Set Filter = SqlTools.ConditionGenerator.Add(TestTab1.id.Name, FieldDataType_Numeric, RelationalOperators_Equal, 2)

   TestTab1.Source.LoadData Filter

   Actual = TestTab1.F1text.Value
   Assert.That Actual, Iz.EqualTo(Expected)

End Sub
