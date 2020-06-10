VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlGeneratorTests_AdodbJet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

'AccUnit:TestClass
'VBA-Inspector:Ignore

Private Generator As ACLibSqlTools.SqlGenerator

' AccUnit infrastructure for advanced AccUnit features. Do not remove these lines.
Implements SimplyVBUnit.ITestFixture
Implements AccUnit_Integration.ITestManagerBridge
Private TestManager As AccUnit_Integration.TestManager
Private Sub ITestManagerBridge_InitTestManager(ByVal NewTestManager As AccUnit_Integration.ITestManagerComInterface): Set TestManager = NewTestManager: End Sub
Private Function ITestManagerBridge_GetTestManager() As AccUnit_Integration.ITestManagerComInterface: Set ITestManagerBridge_GetTestManager = TestManager: End Function
Private Sub ITestFixture_AddTestCases(ByVal Tests As SimplyVBUnit.TestCaseCollector): TestManager.AddTestCases Tests: End Sub

Public Sub Setup()
   Set Generator = SqlTools.SqlGenerator(SqlTools.SqlConverters.JetAdodbSqlConverter)
End Sub

Public Sub Teardown()
   Set Generator = Nothing
End Sub

'--------------------------------------------------------------------
' Tests
'--------------------------------------------------------------------

Public Sub ToString_SelectFields()
   
   Const Expected As String = "Select F1, F2, F3"
   Dim Actual As String
   Dim Fields As ACLibSqlTools.FieldList

   Set Fields = SqlTools.FieldGenerator.FromNames("F1", "F2")

   Actual = Generator.Select(Fields(0), Fields(1), Nothing, "F3").ToString()
   Assert.AreEqual Expected, Actual

End Sub

Public Sub ToString_SelectFieldList()
   
   Const Expected As String = "Select F1, F2, F3"
   Dim Actual As String
   Dim Fields As ACLibSqlTools.FieldList
   Set Fields = SqlTools.FieldGenerator.FromNames("F1", "F2", "F3")

   Actual = Generator.Select(Fields).ToString()
   Assert.AreEqual Expected, Actual

End Sub

Public Sub ToString_From_Select_SpaceInFieldName()
   
   Const Expected As String = "Select [Field 1], Field2 From TableA Where (Field3 = 5)"
   Dim Actual As String
   Actual = Generator.From("TableA").Select("Field 1", "Field2").WhereString("Field3 = 5").ToString()
   Assert.AreEqual Expected, Actual

End Sub

Public Sub ToString_From_Where_GroupBy_Select_SelectField_OrderBy()
   
   Const Expected As String = "Select F1, F2, Count(*) As Cnt From TableA Where (Field3 = 5) Group By F1, F2 Order By F2, F1"
   Dim Actual As String
   Actual = Generator.From("TableA").WhereString("Field3 = 5").GroupBy("F1", "F2") _
                     .Select("F1", "F2").SelectField("Count(*)", "", "Cnt") _
                     .OrderBy("F2", "F1") _
                     .ToString()
   Assert.AreEqual Expected, Actual

End Sub

Public Sub ToJetAdodbString_LikeMitWildcardStern()
   
   Const Expected As String = "Select F1, F2, Count(*) As Cnt From TableA Where (Field3 like 'abc%') Group By F1, F2 Order By F2, F1"
   Dim Actual As String
   Actual = Generator.From("TableA").WhereString("Field3 like 'abc*'").GroupBy("F1", "F2") _
                     .Select("F1", "F2").SelectField("Count(*)", "", "Cnt") _
                     .OrderBy("F2", "F1") _
                     .ToString()
   Assert.AreEqual Expected, Actual

End Sub

Public Sub ToString_WhereWithParams()
   
   Const Expected As String = "Select F1, F2, Count(*) As Cnt From TableA Where (F3 = 5) Group By F1, F2 Order By F2, F1"
   Dim SqlString As String
   SqlString = Generator.From("TableA") _
                        .Where("F3", RelationalOperators.RelationalOperators_Equal, 5) _
                        .GroupBy("F1", "F2") _
                        .Select("F1", "F2").SelectField("Count(*)", "", "Cnt") _
                        .OrderBy("F2", "F1") _
                        .ToString()
   
   Assert.AreEqual Expected, SqlString

End Sub

Public Sub ToString_WhereWithConditionGenerator()
   
   Const Expected As String = "Select * From Tabelle1 Where ((F1 = 1) And ((F2 = 2) Or (F3 = 3)))"

   Dim CondGen As ACLibSqlTools.ConditionGenerator
   
   Set CondGen = SqlTools.ConditionGenerator
   With CondGen
      .Add "F1", FieldDataType_Numeric, RelationalOperators_Equal, 1
      With .BeginGroup(LogicalOperator_Or)
            .Add "F2", FieldDataType_Numeric, RelationalOperators_Equal, 2
            .Add "F3", FieldDataType_Numeric, RelationalOperators_Equal, 3
      End With
   End With
   
   Dim SqlString As String
   SqlString = Generator.From("Tabelle1") _
                        .WhereCondition(CondGen) _
                        .SelectAll _
                        .ToString()
   Assert.AreEqual Expected, SqlString

End Sub
