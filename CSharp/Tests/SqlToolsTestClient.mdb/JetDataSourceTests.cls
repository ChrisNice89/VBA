VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JetDataSourceTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

'AccUnit:TestClass

Private m_DataSource As JetDataSource

' AccUnit infrastructure for advanced AccUnit features. Do not remove these lines.
Implements SimplyVBUnit.ITestFixture
Implements AccUnit_Integration.ITestManagerBridge
Private TestManager As AccUnit_Integration.TestManager
Private Sub ITestManagerBridge_InitTestManager(ByVal NewTestManager As AccUnit_Integration.ITestManagerComInterface): Set TestManager = NewTestManager: End Sub
Private Function ITestManagerBridge_GetTestManager() As AccUnit_Integration.ITestManagerComInterface: Set ITestManagerBridge_GetTestManager = TestManager: End Function
Private Sub ITestFixture_AddTestCases(ByVal Tests As SimplyVBUnit.TestCaseCollector): TestManager.AddTestCases Tests: End Sub

Public Sub Setup()
   Set m_DataSource = New JetDataSource
   m_DataSource.Init "TestTab1", "id", "F1text", "F2long", "F3date"
End Sub

Public Sub Teardown()
   Set m_DataSource = Nothing
End Sub

'--------------------------------------------------------------------
' Tests
'--------------------------------------------------------------------
Public Sub Fields_StateUnderTest_ExpectedBehaviour()
   
   Const Expected As Variant = "id"
   Dim Actual As ACLibSqlTools.Field

   Set Actual = m_DataSource.Fields(1)
   Assert.That Actual.Name, Iz.EqualTo(Expected)

End Sub


Public Sub LoadData_NoFilter_CheckFieldsCount()
   
   Const Expected As Long = 4
   Dim rst As ADODB.Recordset
   Dim Actual As Long

   m_DataSource.LoadData
   Actual = m_DataSource.Fields.Count
   Assert.That Actual, Iz.EqualTo(Expected)

End Sub

Public Sub LoadData_NoFilter_CheckFieldsValue()
   
   Const Expected As Long = 1
   Dim rst As ADODB.Recordset
   Dim Actual As DataField

   m_DataSource.LoadData
   Set Actual = m_DataSource.Fields("id")
   Assert.That Actual.Value, Iz.EqualTo(Expected)

End Sub

Public Sub LoadData_WithFilter_CheckRecordCount()
   
   Const Expected As Long = 1
   Dim rst As ADODB.Recordset
   Dim Actual As Long

   Dim Filter As ACLibSqlTools.ICondition
   Set Filter = SqlTools.ConditionGenerator.Add("id", FieldDataType_Numeric, RelationalOperators_Equal, 1)

   Actual = m_DataSource.LoadData(Filter)
   Assert.That Actual, Iz.EqualTo(Expected)

End Sub
