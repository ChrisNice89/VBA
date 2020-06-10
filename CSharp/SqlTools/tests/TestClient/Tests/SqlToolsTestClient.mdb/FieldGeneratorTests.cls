VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FieldGeneratorTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

Private Generator As ACLibSqlTools.FieldGenerator

'AccUnit:TestClass
'VBA-Inspector:Ignore

' AccUnit infrastructure for advanced AccUnit features. Do not remove these lines.
Implements SimplyVBUnit.ITestFixture
Implements AccUnit_Integration.ITestManagerBridge
Private TestManager As AccUnit_Integration.TestManager
Private Sub ITestManagerBridge_InitTestManager(ByVal NewTestManager As AccUnit_Integration.ITestManagerComInterface): Set TestManager = NewTestManager: End Sub
Private Function ITestManagerBridge_GetTestManager() As AccUnit_Integration.ITestManagerComInterface: Set ITestManagerBridge_GetTestManager = TestManager: End Function
Private Sub ITestFixture_AddTestCases(ByVal Tests As SimplyVBUnit.TestCaseCollector): TestManager.AddTestCases Tests: End Sub

Public Sub Setup()
   Set Generator = SqlTools.FieldGenerator
End Sub

Public Sub Teardown()
   Set Generator = Nothing
End Sub

'--------------------------------------------------------------------
' Tests
'--------------------------------------------------------------------

Public Sub FromNames()

   Dim Fields As ACLibSqlTools.FieldList
   Set Fields = Generator.FromNames("a", "b", "c")

   Assert.That Fields.Item(0).Name, Iz.EqualTo("a")
   Assert.That Fields(1).Name, Iz.EqualTo("b")
   Assert.That Fields(2).Name, Iz.EqualTo("c")

End Sub

Public Sub FromString()

   Dim x() As ACLibSqlTools.Field
   x = Generator.FromString("a, b, c").ToArray

   Assert.That x(0).Name, Iz.EqualTo("a")
   Assert.That x(1).Name, Iz.EqualTo("b")
   Assert.That x(2).Name, Iz.EqualTo("c")

End Sub

Public Sub FromArray_StringArray()

   Dim x() As ACLibSqlTools.Field
   Dim s(2) As String

   s(0) = "a"
   s(1) = "b"
   s(2) = "c"

   x = Generator.FromArray(s).ToArray()

   Assert.That x(0).Name, Iz.EqualTo("a")
   Assert.That x(1).Name, Iz.EqualTo("b")
   Assert.That x(2).Name, Iz.EqualTo("c")

End Sub

Public Sub FromArray_VariantArray()

   Dim x As ACLibSqlTools.FieldList

   Set x = Generator.FromArray(Array("a", "b", "c"))

   Assert.That x(0).Name, Iz.EqualTo("a")
   Assert.That x(1).Name, Iz.EqualTo("b")
   Assert.That x(2).Name, Iz.EqualTo("c")

End Sub

Public Sub FromArray_String_RaiseError()

   Const ArrayTypeMismatchExceptionNumber As Long = -2146233085
   Dim x() As ACLibSqlTools.Field

   Assert.Throws ArrayTypeMismatchExceptionNumber
   x = Generator.FromArray("a, b, c").ToArray

End Sub
