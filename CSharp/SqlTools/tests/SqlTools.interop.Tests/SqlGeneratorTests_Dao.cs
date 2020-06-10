using AccessCodeLib.Data.SqlTools.Sql;
using NUnit.Framework;
using System.Linq;

namespace AccessCodeLib.Data.SqlTools.interop.Tests
{
    [TestFixture]
    class SqlGeneratorTests_Dao : GeneratorTestBase<ISqlGenerator>
    {
        private readonly SqlToolsFactory SqlToolsFactory = new SqlToolsFactory();

        protected override ISqlGenerator GetGenerator()
        {
            return new SqlGenerator(new DaoSqlConverter());
        }

        [Test]
        [TestCase("TableA", "From TableA")]
        [TestCase("Table A", "From [Table A]")]
        public void ToString_From_ToString(string sourceName, string expected)
        {
            var actual = Generator.From(sourceName).ToString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [TestCase("Select Field1", new[] { "Field1" })]
        [TestCase("Select [Field 1]", new[] { "Field 1" })]
        [TestCase("Select Field1, Field2", new[] { "Field1", "Field2" })]
        [TestCase("Select Field1, Field2", new[] { "Field1", "", " ", null, "Field2" })]
        public void ToString_Select_FieldArray_ToString(string expected, string[] fieldNames)
        {
            var fields = SqlToolsFactory.FieldGenerator().FromArray(fieldNames).Cast<IField>().ToArray();
            var actual = Generator.Select(fields).ToString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [TestCase("Select Field1", new[] { "Field1" })]
        [TestCase("Select [Field 1]", new[] { "Field 1" })]
        [TestCase("Select Field1, Field2", new[] { "Field1", "Field2" })]
        [TestCase("Select Field1, Field2", new[] { "Field1", "", " ", null, "Field2" })]
        public void ToString_Select_FieldsList_ToString(string expected, string[] fieldNames)
        {
            var fields = SqlToolsFactory.FieldGenerator().FromArray(fieldNames);
            var actual = Generator.Select(fields).ToString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [TestCase("Select Field1", new [] {"Field1"})]
        [TestCase("Select [Field 1]", new [] {"Field 1"})]
        [TestCase("Select Field1, Field2", new[] { "Field1", "Field2" })]
        [TestCase("Select Field1, Field2", new[] { "Field1", "", " ", null, "Field2" })]
        public void ToString_Select_FieldNames_ToString(string expected, string[] fieldNames)
        {
            var actual = Generator.Select(fieldNames).ToString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [TestCase("Select Field1", "Field1")]
        [TestCase("Select [Field 1]","Field 1")]
        public void ToString_SelectField_ToString(string expected, string fieldName)
        {
            var actual = Generator.SelectField(fieldName, "", "").ToString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void ToString_From_SelectAll_ToString()
        {
            const string expected = "Select * From TableA";
            var actual = Generator.From("TableA").SelectAll().ToString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [TestCase("Group By Field1", new[] { "Field1" })]
        [TestCase("Group By [Field 1]", new[] { "Field 1" })]
        [TestCase("Group By Field1, Field2", new[] { "Field1", "Field2" })]
        [TestCase("Group By Field1, Field2", new[] { "Field1", "", " ", null, "Field2" })]
        public void ToString_GroupBy_Fields_ToString(string expected, string[] fieldNames)
        {
            var fields = SqlToolsFactory.FieldGenerator().FromArray(fieldNames);
            var actual = Generator.GroupBy(fields).ToString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [TestCase("Order By Field1", new[] { "Field1" })]
        [TestCase("Order By Field1, Field2", new[] { "Field1", "Field2" })]
        public void ToString_OrderBy_Fields_ToString(string expected, string[] fieldNames)
        {
            var fields = SqlToolsFactory.FieldGenerator().FromArray(fieldNames).Cast<IField>().ToArray();
            var actual = Generator.OrderBy(fields).ToString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void ToString_From_Select_Select_Where_ToString()
        {
            const string expected = "Select Field1, Field2 From TableA Where (Field3 = 3)";
            var actual = Generator.From("TableA").SelectField("Field1", "", "").Select("Field2").Where("Field3 = 3").ToString();

            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void ToString_From_Select2Fields_Where_ToString()
        {
            const string expected = "Select Field1, Field2 From TableA Where (Field3 = 3)";
            var actual = Generator.From("TableA").Select("Field1", "Field2").Where("Field3 = 3").ToString();

            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void ToString_From_Select2Fields_Where_GroupBy_Having_OrderBy_ToString()
        {
            const string expected = "Select Field1, Field2, Count(*) As Cnt From TableA Where (Field3 = 3) Group By Field1, Field2 Having (Count(*) > 2) Order By Field1, Field2";
            var actual = Generator.From("TableA").Select("Field1", "Field2").SelectField("Count(*)", "", "Cnt").Where("Field3 = 3").GroupBy("Field1", "Field2").Having("Count(*) > 2").OrderBy("Field1", "Field2").ToString();

            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void Where_WithOneCondition_ToString()
        {
            const string expected = "Where F1 = 123";

            var actual = Generator.Where(new Field("F1"), RelationalOperators.Equal, 123)
                                      .ToString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void Where_Between_ToString()
        {
            const string expected = "Where (F1 Between 1 And 9)";

            var actual = Generator.WhereBetween(new Field("F1"), 1, 9).ToString();

            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void Where_WithTwoConditionCombinedWithAnd_ToString()
        {
            const string expected = "Where F1 = 123 And F2 = 456";

            var actual = Generator.Where(new Field("F1"), RelationalOperators.Equal, 123)
                                      .Where(new Field("F2"), RelationalOperators.Equal, 456)
                                      .ToString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void Where_WithFourConditionsCombinedWithOrAndAnd1_ToString()
        {
            const string expected = "Where ((F1 = 123 And F2 = 456) Or (F1 = 789 And F2 = 98))";

            var conditionGenerator = new ConditionGenerator {ConcatOperator = LogicalOperator.Or};

            conditionGenerator.BeginGroup(new Field("F1"), RelationalOperators.Equal, 123)
                              .Add(new Field("F2"), FieldDataType.Numeric, RelationalOperators.Equal, 456)
                              .ConcatOperator = LogicalOperator.And;

            conditionGenerator.BeginGroup(new Field("F1"), RelationalOperators.Equal, 789)
                              .Add(new Field("F2"), FieldDataType.Numeric, RelationalOperators.Equal, 98)
                              .ConcatOperator = LogicalOperator.And;

            var actual = Generator.Where(conditionGenerator).ToString();

            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void Where_WithFourConditionsCombinedWithOrAndAnd2_ToString()
        {
            const string expected = "Where ((F1 = 123 Or F2 = 456) And (F1 = 789 Or F2 = 98))";

            var conditionGenerator = new ConditionGenerator();

            conditionGenerator.BeginGroup(new Field("F1"), RelationalOperators.Equal, 123)
                              .Add(new Field("F2"), FieldDataType.Numeric, RelationalOperators.Equal, 456)
                              .ConcatOperator = LogicalOperator.Or;

            conditionGenerator.BeginGroup(new Field("F1"), RelationalOperators.Equal, 789)
                              .Add(new Field("F2"), FieldDataType.Numeric, RelationalOperators.Equal, 98)
                              .ConcatOperator = LogicalOperator.Or;

            var actual = Generator.Where(conditionGenerator).ToString();

            Assert.AreEqual(expected, actual);
        }

        
        [Test]
        public void Where_WithStringValue_ToString()
        {
            const string expected = "Where F1 = '123'";

            var actual = Generator.Where(new Field("F1"), RelationalOperators.Equal, "123")
                                      .ToString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void Where_WithTwoFieldAndStringValues_ToString()
        {
            const string expected = "Where F1 = '123' And F2 >= 'abc'";

            var actual = Generator.Where(new Field("F1"), RelationalOperators.Equal, "123")
                                      .Where(new Field("F2"), RelationalOperators.GreaterThan | RelationalOperators.Equal, "abc")
                                      .ToString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void Where_WithField_ToString()
        {
            const string expected = "Where F1 = F2";

            var actual = Generator.Where(new Field("F1"), RelationalOperators.Equal, new Field("F2"))
                                      .ToString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void From_NamedSource_ToString()
        {
            const string expected = "From Tab1";

            var actual = Generator.From(new NamedSource("Tab1"))
                                      .ToString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void From_NamedSourceWithAlias_ToString()
        {
            const string expected = "From Tab1"; // DAO/Jet/ACE kennt kein Schema!

            var actual = Generator.From(new NamedSource("Tab1", "dbo"))
                                      .ToString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void From_Subselect_ToString()
        {
            const string expected = "From (Select F1 From Tab1)";

            var subSelectGenerator = new SqlGenerator();
            var subSelect = new SubSelectSource(subSelectGenerator.Select("F1").From("Tab1").SqlStatement);

            var actual = Generator.From(subSelect)
                                      .ToString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void Having_WithOneCondition_ToString()
        {
            const string expected = "Having F1 = 123";

            var actual = Generator.Having(new Field("F1"), RelationalOperators.Equal, 123)
                                      .ToString();
            Assert.AreEqual(expected, actual);
        }
    }
}
