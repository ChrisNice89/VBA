using AccessCodeLib.Data.Common.Sql;
using AccessCodeLib.Data.SqlTools;
using AccessCodeLib.Data.SqlTools.Sql;
using NUnit.Framework;

namespace AccessCodeLib.data.SqlTools.Tests
{
    [TestFixture]
    public class SqlGeneratorTestsTsql
    {
        private ISqlGenerator _sqlGenerator;

        [SetUp]
        public void MyTestInitialize()
        {
            var converter = new Data.SqlTools.Converter.TSQL.SqlConverter();
            _sqlGenerator = new SqlGenerator(converter);
        }

        [TearDown]
        public void MyTestCleanup()
        {
            _sqlGenerator = null;
        }

        [Test]
        [TestCase("TableA", "From TableA")]
        [TestCase("Table A", "From \"Table A\"")]
        public void ToString_From_CheckString(string sourceName, string expected)
        {
            var actual = _sqlGenerator.From(sourceName).ToString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [TestCase("Select Field1", new[] { "Field1" })]
        [TestCase("Select \"Field 1\"", new[] { "Field 1" })]
        [TestCase("Select Field1, Field2", new[] { "Field1", "Field2" })]
        [TestCase("Select Field1, Field2", new[] { "Field1", "", " ", null, "Field2" })]
        public void ToString_Select_Fields_CheckString(string expected, string[] fieldNames)
        {
            var fields = SqlToolsFactory.FieldGenerator.FromArray(fieldNames);
            var actual = _sqlGenerator.Select(fields).ToString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [TestCase("Select Field1", new[] { "Field1" })]
        [TestCase("Select \"Field 1\"", new[] { "Field 1" })]
        [TestCase("Select Field1, Field2", new[] { "Field1", "Field2" })]
        [TestCase("Select Field1, Field2", new[] { "Field1", "", " ", null, "Field2" })]
        public void ToString_Select_FieldNames_CheckString(string expected, string[] fieldNames)
        {
            var actual = _sqlGenerator.Select(fieldNames).ToString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [TestCase("Select Field1", "Field1")]
        [TestCase("Select \"Field 1\"", "Field 1")]
        public void ToString_SelectField_CheckString(string expected, string fieldName)
        {
            var actual = _sqlGenerator.SelectField(fieldName).ToString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void ToString_From_SelectAll_CheckString()
        {
            const string expected = "Select * From TableA";
            var actual = _sqlGenerator.From("TableA").SelectAll().ToString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [TestCase("Group By Field1", new[] { "Field1" })]
        [TestCase("Group By \"Field 1\"", new[] { "Field 1" })]
        [TestCase("Group By Field1, Field2", new[] { "Field1", "Field2" })]
        [TestCase("Group By Field1, Field2", new[] { "Field1", "", " ", null, "Field2" })]
        public void ToString_GroupBy_Fields_CheckString(string expected, string[] fieldNames)
        {
            var fields = SqlToolsFactory.FieldGenerator.FromArray(fieldNames);
            var actual = _sqlGenerator.GroupBy(fields).ToString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [TestCase("Order By Field1", new[] { "Field1" })]
        [TestCase("Order By Field1, Field2", new[] { "Field1", "Field2" })]
        public void ToString_OrderBy_Fields_CheckString(string expected, string[] fieldNames)
        {
            var fields = SqlToolsFactory.FieldGenerator.FromArray(fieldNames);
            var actual = _sqlGenerator.OrderBy(fields).ToString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void ToString_From_Select_Select_Where_CheckString()
        {
            const string expected = "Select Field1, Field2 From TableA Where (Field3 = 3)";
            var actual = _sqlGenerator.From("TableA").SelectField("Field1").SelectField("Field2").WhereString("Field3 = 3").ToString();

            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void ToString_From_Select2Fields_Where_CheckString()
        {
            const string expected = "Select Field1, Field2 From TableA Where (Field3 = 3)";
            var actual = _sqlGenerator.From("TableA").Select("Field1", "Field2").WhereString("Field3 = 3").ToString();

            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void ToString_From_Select2Fields_Where_GroupBy_Having_OrderBy_CheckString()
        {
            const string expected = "Select Field1, Field2, Count(*) As Cnt From TableA Where (Field3 = 3) Group By Field1, Field2 Having (Count(*) > 2) Order By Field1, Field2";
            var actual = _sqlGenerator.From("TableA").Select("Field1", "Field2").SelectField("Count(*)", null, "Cnt").WhereString("Field3 = 3").GroupBy("Field1", "Field2").Having("Count(*) > 2").OrderBy("Field1", "Field2").ToString();

            Assert.AreEqual(expected, actual);
        }

        [Test]
        [TestCase("Where (Field3 like 'a%')", "Field3 like 'a*'")]
        [TestCase("Where (Field3 like 'a[*]')", "Field3 like 'a[*]'")]
        [TestCase("Where ([F*3] like 'a%')", "[F*3] like 'a*'")]
        [TestCase("Where (Field3 like 'a[xy*]')", "Field3 like 'a[xy*]'")]
        [TestCase("Where (Field3 like 'a[xy*]' and F4 like 'a%b')", "Field3 like 'a[xy*]' and F4 like 'a*b'")]
        [TestCase("Where ([F_1] like 'a_' and F2 like 'a[_]b' and F3 like 'a[%]b')", "[F_1] like 'a?' and F2 like 'a_b' and F3 like 'a%b'")]
        public void ToString_Where_CheckString(string expected, string whereString)
        {
            var actual = _sqlGenerator.WhereString(whereString).ToString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void ToString_From_Select2Fields_Where_GroupBy_Having_OrderBy()
        {
            const string expected = "Select Field1, Field2, Count(*) As Cnt From TableA Where (Field3 = 3 And Field4 like 'a%') Group By Field1, Field2 Having (Count(*) > 2) Order By Field1, Field2";
            var actual =
                _sqlGenerator.From("TableA").Select("Field1", "Field2")
                                .SelectField("Count(*)", null, "Cnt")
                                .WhereString("Field3 = 3 And Field4 like 'a*'")
                                .GroupBy("Field1", "Field2")
                                .Having("Count(*) > 2")
                                .OrderBy("Field1", "Field2")
                                .ToString();

            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void Where_WithOneCondition_CheckString()
        {
            const string expected = "Where (F1 = 123)";

            var actual = _sqlGenerator.Where(new Field("F1"), RelationalOperators.Equal, 123)
                                        .ToString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void Where_WithTwoConditionCombinedWithAnd_CheckString()
        {
            const string expected = "Where (F1 = 123) And (F2 = 456)";

            var actual = _sqlGenerator.Where(new Field("F1"), RelationalOperators.Equal, 123)
                                        .Where(new Field("F2"), RelationalOperators.Equal, 456)
                                        .ToString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void Where_WithFourConditionsCombinedWithOrAndAnd1_CheckString()
        {
            const string expected = "Where (((F1 = 123) And (F2 = 456)) Or ((F1 = 789) And (F2 = 98)))";

            var conditionGenerator = new ConditionGenerator { ConcatOperator = LogicalOperator.Or };

            conditionGenerator.BeginGroup(new Field("F1"), RelationalOperators.Equal, 123)
                                .Add(new Field("F2"), RelationalOperators.Equal, 456)
                                .ConcatOperator = LogicalOperator.And;

            conditionGenerator.BeginGroup(new Field("F1"), RelationalOperators.Equal, 789)
                                .Add(new Field("F2"), RelationalOperators.Equal, 98)
                                .ConcatOperator = LogicalOperator.And;

            var actual = _sqlGenerator.Where(conditionGenerator).ToString();

            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void Where_WithFourConditionsCombinedWithOrAndAnd2_CheckString()
        {
            const string expected = "Where (((F1 = 123) Or (F2 = 456)) And ((F1 = 789) Or (F2 = 98)))";

            var conditionGenerator = new ConditionGenerator();

            conditionGenerator.BeginGroup(new Field("F1"), RelationalOperators.Equal, 123)
                                .Add(new Field("F2"), RelationalOperators.Equal, 456)
                                .ConcatOperator = LogicalOperator.Or;

            conditionGenerator.BeginGroup(new Field("F1"), RelationalOperators.Equal, 789)
                                .Add(new Field("F2"), RelationalOperators.Equal, 98)
                                .ConcatOperator = LogicalOperator.Or;

            var actual = _sqlGenerator.Where(conditionGenerator).ToString();

            Assert.AreEqual(expected, actual);
        }


        [Test]
        public void Where_WithStringValue_CheckString()
        {
            const string expected = "Where (F1 = '123')";

            var actual = _sqlGenerator.Where(new Field("F1"), RelationalOperators.Equal, "123")
                                        .ToString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void Where_WithTwoFieldAndStringValues_CheckString()
        {
            const string expected = "Where (F1 = '123') And (F2 >= 'abc')";

            var actual = _sqlGenerator.Where(new Field("F1"), RelationalOperators.Equal, "123")
                                        .Where(new Field("F2"), RelationalOperators.GreaterThan | RelationalOperators.Equal, "abc")
                                        .ToString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void Where_WithField_CheckString()
        {
            const string expected = "Where F1 = F2";

            var actual = _sqlGenerator.Where(new Field("F1"), RelationalOperators.Equal, new Field("F2"))
                                        .ToString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void From_NamedSource_ToString()
        {
            const string expected = "From Tab1";

            var actual = _sqlGenerator.From(new NamedSource("Tab1"))
                                      .ToString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void From_NamedSourceWithAlias_ToString()
        {
            const string expected = "From dbo.Tab1";

            var actual = _sqlGenerator.From(new NamedSource("Tab1", "dbo"))
                                      .ToString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void Select_WithSource_ToString()
        {
            const string expected = "Select Tab1.F1, Tab1.F2";

            var source = new NamedSource("Tab1");
            var actual = _sqlGenerator.Select(new Field("F1", source), new Field("F2", source))
                                      .ToString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void Select_WithSourceWithSchema_ToString()
        {
            const string expected = "Select dbo.Tab1.F1, dbo.Tab1.F2";

            var source = new NamedSource("Tab1", "dbo");
            var actual = _sqlGenerator.Select(new Field("F1", source), new Field("F2", source))
                                      .ToString();
            Assert.AreEqual(expected, actual);
        }
    }
}
