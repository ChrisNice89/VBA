using System;
using AccessCodeLib.Data.Common.Sql;
using AccessCodeLib.Data.SqlTools.Converter.Jet;
using AccessCodeLib.Data.SqlTools.Sql;
using NUnit.Framework;

namespace AccessCodeLib.Data.SqlTools.Converter.Tests.Jet.Dao
{
// ReSharper disable UnusedMember.Global
    class ConditionConverterTests : ConditionConverterTestBase
// ReSharper restore UnusedMember.Global
    {
        protected override IConditionConverter GetConverter()
        {
            return new ConditionConverter(new NameConverter(), new Converter.Jet.ValueConverter());
        }

        [Test]
        public void ConditionFieldWithNamesSource()
        {
            const string expected = "Tab.F1 = 1";
            Conditions.Add(new Field("F1", new NamedSource("Tab")), RelationalOperators.Equal, 1);

            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void ConditionFieldWithNamesSource_WithSchema()
        {
            const string expected = "Tab.F1 = 1"; // kein Schema, da Dao
            Conditions.Add(new Field("F1", new NamedSource("Tab", "dbo")), RelationalOperators.Equal, 1);

            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void ConditionFieldWithNamesSource_WithAlias()
        {
            const string expected = "X.F1 = 1"; // kein Schema, da Dao
            Conditions.Add(new Field("F1", new SourceAlias(new NamedSource("Tab"), "X")), RelationalOperators.Equal, 1);

            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void ConditionFieldWithSubSelectSource_WithAlias()
        {
            const string expected = "X.F1 = 1";

            var sqlGenerator = new SqlGenerator();
            var subSelect = sqlGenerator.From("Tab1").Select("F1").SqlStatement;

            var source = new SourceAlias(new SubSelectSource(subSelect), "X");

            Conditions.Add(new Field("F1", source), RelationalOperators.Equal, 1);

            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void ConditionFieldWithSubSelectSource_MissingAlias()
        {
            const string expected = "F1 = 1";

            var sqlGenerator = new SqlGenerator();
            var subSelect = sqlGenerator.From("Tab1").Select("F1").SqlStatement;

            var source = new SubSelectSource(subSelect);

            Conditions.Add(new Field("F1", source), RelationalOperators.Equal, 1);

            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void ConditionField_DoubleValue_MissingAlias()
        {
            const string expected = "F1 = 1.2";

            Conditions.Add(new Field("F1"), RelationalOperators.Equal, 1.2);

            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void ConditionField_Like_StringWithWildcard()
        {
            const string expected = "F1 Like 'abc*'";

            Conditions.Add(new Field("F1"), RelationalOperators.Like, "abc*");

            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [ExpectedException(typeof(NotSupportedRelationalOperatorException))]
        public void ConditionField_LikeWithEqual_ThrowException()
        {
            Conditions.Add(new Field("F1"), RelationalOperators.Like | RelationalOperators.Equal, "abc*");
            GenerateSqlString();
        }

        [Test]
        public void ConditionField_Between_IntValues()
        {
            const string expected = "(F1 Between 1 And 5)";

            Conditions.Add(new Field("F1"), RelationalOperators.Between, new BetweenValue(1, 5));

            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void ConditionField_Between_DoubleValues()
        {
            const string expected = "(F1 Between 1.2 And 3.4)";

            Conditions.Add(new Field("F1"), RelationalOperators.Between, new BetweenValue(new NumericValue<double>(1.2), new NumericValue<double>(3.4)));

            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void ConditionField_Between_DecimalValues()
        {
            const string expected = "(F1 Between 1.2 And 3.4)";

            Conditions.Add(new Field("F1"), RelationalOperators.Between, new BetweenValue(new NumericValue<decimal>(1.2m), new NumericValue<decimal>(3.4m)));

            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void ConditionField_Between_StringValues()
        {
            const string expected = "(F1 Between 'a' And 'c')";

            Conditions.Add(new Field("F1"), RelationalOperators.Between, new BetweenValue(new TextValue("a"), new TextValue("c")));

            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void ConditionField_Between_IntValuesWithDbNull()
        {
            const string expected = "F1 >= 1";

            IValue FirstValue = new NumericValue<int>(1);
            IValue SecondValue = new NullValue();

            Conditions.Add(new Field("F1"), RelationalOperators.Between, new BetweenValue(FirstValue, SecondValue));

            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void ConditionField_In_NumericValues()
        {
            const string expected = "F1 In (0,1,2.3)";

            var values = new double[] { 0, 1, 2.3 };
            var iValues = new IValue[values.Length];
            for (var i = 0; i < values.Length; i++)
            {
                iValues[i] = new NumericValue<double>(values[i]);
            }

            Conditions.Add(new Field("F1"), RelationalOperators.In, new ValueArray(iValues));

            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void ConditionField_In_StringValues()
        {
            const string expected = "F1 In ('a','b','c')";
            
            var values = new string[] { "a", "b", "c" };
            var iValues = new IValue[values.Length];
            for (var i = 0; i < values.Length; i++)
            {
               iValues[i] = new TextValue(values[i]);
            }

            Conditions.Add(new Field("F1"), RelationalOperators.In, new ValueArray(iValues));

            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void ConditionField_In_DateTimeValues()
        {
            const string expected = "F1 In (#2015-01-01#,#2015-02-01#)";

            var values = new DateTime[] { System.DateTime.Parse("2015-01-01"), System.DateTime.Parse("2015-02-01") };
            var iValues = new IValue[values.Length];
            for (var i = 0; i < values.Length; i++)
            {
                iValues[i] = new DateTimeValue(values[i]);
            }

            Conditions.Add(new Field("F1"), RelationalOperators.In, new ValueArray(iValues));

            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }
    }
}
