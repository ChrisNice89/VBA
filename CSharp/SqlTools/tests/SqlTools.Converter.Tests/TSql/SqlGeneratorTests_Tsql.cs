using AccessCodeLib.Data.Common.Sql;
using AccessCodeLib.Data.Common.Sql.Converter;
using AccessCodeLib.Data.SqlTools.Sql;
using NUnit.Framework;

namespace AccessCodeLib.Data.SqlTools.Converter.Tests.TSql
{
    [TestFixture]
    class SqlConverterTests : SqlConverterTestBase
    {
        protected override ISqlConverter GetConverter()
        {
            return new Mssql.SqlConverter();
        }

        [Test]
        [TestCase("TableA", "From TableA")]
        [TestCase("Table A", "From \"Table A\"")]
        public void ToString_From_CheckString(string sourceName, string expected)
        {
            Generator.From(sourceName);
            var actual = GenerateSqlString();
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
            Generator.Select(fields);
            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [TestCase("Select Field1", new[] { "Field1" })]
        [TestCase("Select \"Field 1\"", new[] { "Field 1" })]
        [TestCase("Select Field1, Field2", new[] { "Field1", "Field2" })]
        [TestCase("Select Field1, Field2", new[] { "Field1", "", " ", null, "Field2" })]
        public void ToString_Select_FieldNames_CheckString(string expected, string[] fieldNames)
        {
            Generator.Select(fieldNames);
            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [TestCase("Select Field1", "Field1")]
        [TestCase("Select \"Field 1\"", "Field 1")]
        public void ToString_SelectField_CheckString(string expected, string fieldName)
        {
            Generator.SelectField(fieldName);
            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void ToString_From_SelectAll_CheckString()
        {
            const string expected = "Select * From TableA";
            Generator.From("TableA").SelectAll();
            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void ToString_FromAlias_SelectAll_CheckString()
        {
            const string expected = "Select * From TableA As X";
            Generator.From(new SourceAlias(new NamedSource("TableA"), "X")).SelectAll();
            var actual = GenerateSqlString();
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
            Generator.GroupBy(fields);
            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [TestCase("Order By Field1", new[] { "Field1" })]
        [TestCase("Order By Field1, Field2", new[] { "Field1", "Field2" })]
        public void ToString_OrderBy_Fields_CheckString(string expected, string[] fieldNames)
        {
            var fields = SqlToolsFactory.FieldGenerator.FromArray(fieldNames);
            Generator.OrderBy(fields);
            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void ToString_From_Select_Select_Where_CheckString()
        {
            const string expected = "Select Field1, Field2 From TableA Where (Field3 = 3)";
            Generator.From("TableA").SelectField("Field1").SelectField("Field2").Where("Field3 = 3");
            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void ToString_From_Select2Fields_Where_CheckString()
        {
            const string expected = "Select Field1, Field2 From TableA Where (Field3 = 3)";
            Generator.From("TableA").Select("Field1", "Field2").Where("Field3 = 3");
            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void ToString_From_Select2Fields_Where_GroupBy_Having_OrderBy_CheckString()
        {
            const string expected = "Select Field1, Field2, Count(*) As Cnt From TableA Where (Field3 = 3) Group By Field1, Field2 Having (Count(*) > 2) Order By Field1, Field2";
            Generator.From("TableA").Select("Field1", "Field2").SelectField("Count(*)", null, "Cnt").Where("Field3 = 3").GroupBy("Field1", "Field2").Having("Count(*) > 2").OrderBy("Field1", "Field2");
            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [TestCase("Where (Field3 Like 'a%')", "Field3 Like 'a*'")]
        [TestCase("Where (Field3 Like 'a[*]')", "Field3 Like 'a[*]'")]
        [TestCase("Where ([F*3] Like 'a%')", "[F*3] Like 'a*'")]
        [TestCase("Where (Field3 Like 'a[xy*]')", "Field3 Like 'a[xy*]'")]
        [TestCase("Where (Field3 Like 'a[xy*]' and F4 Like 'a%b')", "Field3 Like 'a[xy*]' and F4 Like 'a*b'")]
        [TestCase("Where ([F_1] Like 'a_' and F2 Like 'a[_]b' and F3 Like 'a[%]b')", "[F_1] Like 'a?' and F2 Like 'a_b' and F3 Like 'a%b'")]
        public void ToString_Where_CheckString(string expected, string whereString)
        {
            Generator.Where(whereString);
            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void ToString_From_Select2Fields_Where_GroupBy_Having_OrderBy()
        {
            const string expected = "Select Field1, Field2, Count(*) As Cnt From TableA Where (Field3 = 3 And Field4 Like 'a%') Group By Field1, Field2 Having (Count(*) > 2) Order By Field1, Field2";
            
            Generator.From("TableA").Select("Field1", "Field2")
                     .SelectField("Count(*)", null, "Cnt")
                     .Where("Field3 = 3 And Field4 Like 'a*'")
                     .GroupBy("Field1", "Field2")
                     .Having("Count(*) > 2")
                     .OrderBy("Field1", "Field2");

            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void Where_WithOneCondition_CheckString()
        {
            const string expected = "Where F1 = 123";
            Generator.Where(new Field("F1"), RelationalOperators.Equal, 123);
            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void Where_WithTwoConditionCombinedWithAnd_CheckString()
        {
            const string expected = "Where F1 = 123 And F2 = 456";

            Generator.Where(new Field("F1"), RelationalOperators.Equal, 123)
                     .Where(new Field("F2"), RelationalOperators.Equal, 456);

            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void Where_WithFourConditionsCombinedWithOrAndAnd1_CheckString()
        {
            const string expected = "Where ((F1 = 123 And F2 = 456) Or (F1 = 789 And F2 = 98))";

            var conditionGenerator = new ConditionGenerator { ConcatOperator = LogicalOperator.Or };

            conditionGenerator.BeginGroup()
                                .Add(new Field("F1"), RelationalOperators.Equal, 123)
                                .Add(new Field("F2"), RelationalOperators.Equal, 456)
                                .ConcatOperator = LogicalOperator.And;

            conditionGenerator.BeginGroup()
                                .Add(new Field("F1"), RelationalOperators.Equal, 789)
                                .Add(new Field("F2"), RelationalOperators.Equal, 98)
                                .ConcatOperator = LogicalOperator.And;

            Generator.Where(conditionGenerator);

            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void Where_WithFourConditionsCombinedWithOrAndAnd2_CheckString()
        {
            const string expected = "Where ((F1 = 123 Or F2 = 456) And (F1 = 789 Or F2 = 98))";

            var conditionGenerator = new ConditionGenerator();

            conditionGenerator.BeginGroup()
                                .Add(new Field("F1"), RelationalOperators.Equal, 123)
                                .Add(new Field("F2"), RelationalOperators.Equal, 456)
                                .ConcatOperator = LogicalOperator.Or;

            conditionGenerator.BeginGroup()
                                .Add(new Field("F1"), RelationalOperators.Equal, 789)
                                .Add(new Field("F2"), RelationalOperators.Equal, 98)
                                .ConcatOperator = LogicalOperator.Or;

            Generator.Where(conditionGenerator);

            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }


        [Test]
        public void Where_WithStringValue_CheckString()
        {
            const string expected = "Where F1 = '123'";
            Generator.Where(new Field("F1"), RelationalOperators.Equal, "123");
            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void Where_WithTwoFieldAndStringValues_CheckString()
        {
            const string expected = "Where F1 = '123' And F2 >= 'abc'";

            Generator.Where(new Field("F1"), RelationalOperators.Equal, "123")
                     .Where(new Field("F2"), RelationalOperators.GreaterThan | RelationalOperators.Equal, "abc");

            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void Where_WithField_CheckString()
        {
            const string expected = "Where F1 = F2";
            Generator.Where(new Field("F1"), RelationalOperators.Equal, new Field("F2"));
            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void From_NamedSource_ToString()
        {
            const string expected = "From Tab1";
            Generator.From(new NamedSource("Tab1"));
            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void From_NamedSourceWithAlias_ToString()
        {
            const string expected = "From dbo.Tab1";
            Generator.From(new NamedSource("Tab1", "dbo"));
            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void Select_WithSource_ToString()
        {
            const string expected = "Select Tab1.F1, Tab1.F2";

            var source = new NamedSource("Tab1");
            Generator.Select(new Field("F1", source), new Field("F2", source));

            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void Select_WithSourceWithSchema_ToString()
        {
            const string expected = "Select dbo.Tab1.F1, dbo.Tab1.F2";

            var source = new NamedSource("Tab1", "dbo");
            Generator.Select(new Field("F1", source), new Field("F2", source));

            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }
    }
}
