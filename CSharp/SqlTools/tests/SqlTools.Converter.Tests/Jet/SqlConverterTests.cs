using AccessCodeLib.Data.Common.Sql;
using AccessCodeLib.Data.SqlTools.Sql;
using NUnit.Framework;

namespace AccessCodeLib.Data.SqlTools.Converter.Tests.Jet
{
    //
    // Allgemeine Test-Einstellungen, die für DAO und ADODB gelten
    // Beispiel: Tests für Tabellen- u. Feldnamen
    //
    abstract class SqlConverterTests : SqlConverterTestBase
    {  
        // die folgenden Tests werden weitervererbt und dann je TestFixture getestet

        [Test]
        [TestCase("TableA", "From TableA")]
        [TestCase("Table A", "From [Table A]")]
        public void From_SourceAsString_1Source(string sourceName, string expected)
        {
            Generator.From(sourceName);
            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [TestCase("Select Field1", new[] { "Field1" })]
        [TestCase("Select [Field 1]", new[] { "Field 1" })]
        [TestCase("Select Field1, Field2", new[] { "Field1", "Field2" })]
        [TestCase("Select Field2, Field3", new[] { "Field2", "", " ", null, "Field3" })]
        public void Select_Fields_FromFieldGenerator(string expected, string[] fieldNames)
        {
            var fields = SqlToolsFactory.FieldGenerator.FromArray(fieldNames);
            Generator.Select(fields);
            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [TestCase("Select Field1", new[] { "Field1" })]
        [TestCase("Select [Field 1]", new[] { "Field 1" })]
        [TestCase("Select Field1, Field2", new[] { "Field1", "Field2" })]
        [TestCase("Select Field1, Field3", new[] { "Field1", "", " ", null, "Field3" })]
        public void Select_FieldNames_StringArray(string expected, string[] fieldNames)
        {
            Generator.Select(fieldNames);
            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [TestCase("Select Field1 As F1", "Field1", "F1")]
        [TestCase("Select [Field 1] As F1", "Field 1", "F1")]
        public void Select_FieldNameWithAlias(string expected, string fieldName, string alias)
        {
            Generator.Select(new FieldAlias(new Field(fieldName), alias));
            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [TestCase("Select Field1", "Field1")]
        [TestCase("Select [Field 1]", "Field 1")]
        public void Select_Field_FieldName(string expected, string fieldName)
        {
            Generator.Select(fieldName);
            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void SelectAll_From_SourceString()
        {
            const string expected = "Select * From TableA";
            Generator.From("TableA").SelectAll();
            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [TestCase("Group By Field1", new[] { "Field1" })]
        [TestCase("Group By [Field 1]", new[] { "Field 1" })]
        [TestCase("Group By Field1, Field2", new[] { "Field1", "Field2" })]
        [TestCase("Group By Field1, Field2", new[] { "Field1", "", " ", null, "Field2" })]
        public void GroupBy_Fields_StringArray(string expected, string[] fieldNames)
        {
            var fields = SqlToolsFactory.FieldGenerator.FromArray(fieldNames);
            Generator.GroupBy(fields);
            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [TestCase("Order By Field1", new[] { "Field1" })]
        [TestCase("Order By Field1, Field2", new[] { "Field1", "Field2" })]
        public void OrderBy_Fields_StringArray(string expected, string[] fieldNames)
        {
            var fields = SqlToolsFactory.FieldGenerator.FromArray(fieldNames);
            Generator.OrderBy(fields);
            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void From_Select_Select_Where()
        {
            const string expected = "Select Field1, Field2 From TableA Where (Field3 = 3)";
            Generator.From("TableA").SelectField("Field1").SelectField("Field2").Where("Field3 = 3");

            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void From_Select2Fields_Where()
        {
            const string expected = "Select Field1, Field2 From TableA Where (Field3 = 3)";
            Generator.From("TableA").Select("Field1", "Field2").Where("Field3 = 3");
            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void From_Select2Fields_Where_GroupBy_Having_OrderBy()
        {
            const string expected = "Select Field1, Field2, Count(*) As Cnt From TableA Where (Field3 = 3) Group By Field1, Field2 Having (Count(*) > 2) Order By Field1, Field2";
            Generator.From("TableA").Select("Field1", "Field2")
                      .SelectField("Count(*)", null, "Cnt")
                     .Where("Field3 = 3")
                     .GroupBy("Field1", "Field2")
                     .Having("Count(*) > 2")
                     .OrderBy("Field1", "Field2");
            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void WithOneCondition()
        {
            const string expected = "Where F1 = 123";
            Generator.Where(new Field("F1"), RelationalOperators.Equal, 123);
            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void Where_WithTwoConditionCombinedWithAnd()
        {
            const string expected = "Where F1 = 123 And F2 = 456";

            Generator.Where(new Field("F1"), RelationalOperators.Equal, 123)
                     .Where(new Field("F2"), RelationalOperators.Equal, 456);

            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void Where_WithFourConditionsCombinedWithOrAndAnd1()
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
        public void Where_WithFourConditionsCombinedWithOrAndAnd2()
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
        public void Where_WithStringValue()
        {
            const string expected = "Where F1 = '123'";

            Generator.Where(new Field("F1"), RelationalOperators.Equal, "123");

            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void Where_WithStringValue_Like()
        {
            const string expected = "Where F1 Like '123'";

            Generator.Where(new Field("F1"), RelationalOperators.Like, "123");

            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void Where_WithTwoFieldAndStringValues_ToString()
        {
            const string expected = "Where F1 = '123' And F2 >= 'abc'";

            Generator.Where(new Field("F1"), RelationalOperators.Equal, "123")
                     .Where(new Field("F2"), RelationalOperators.GreaterThan | RelationalOperators.Equal, "abc");

            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void Where_WithField_ToString()
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
            const string expected = "From Tab1"; // DAO/Jet/ACE kennt kein Schema!
            Generator.From(new NamedSource("Tab1", "dbo"));
            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void From_Subselect_ToString()
        {
            const string expected = "From (Select F1 From Tab1)";

            var subSelectGenerator = new SqlGenerator();
            var subSelect = new SubSelectSource(subSelectGenerator.Select("F1").From("Tab1").SqlStatement);

            Generator.From(subSelect);

            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void From_SubselectWithAlias()
        {
            const string expected = "From (Select F1 From Tab1) As X";

            var subSelectGenerator = new SqlGenerator();
            var subSelect = new SubSelectSource(subSelectGenerator.Select("F1").From("Tab1").SqlStatement);

            Generator.From(new SourceAlias(subSelect, "X"));

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
            const string expected = "Select Tab1.F1, Tab1.F2";

            var source = new NamedSource("Tab1", "dbo");
            Generator.Select(new Field("F1", source), new Field("F2", source));

            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Ignore("ExistsCondition fehlt noch")]
        public void WhereExists_ToString()
        {
            const string expected = "Where Exists(select * from Tab2 Where X = 0)";
            ICondition condition = null; // new ExistsCondition();
            Generator.Where(condition);
            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void Having_Field_StringValue()
        {
            const string expected = "Having F1 = '123'";

            Generator.Having(new Field("F1"), RelationalOperators.Equal, "123");

            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }
    }
}
