using NUnit.Framework;

namespace AccessCodeLib.Data.SqlTools.interop.Tests
{
    [TestFixture]
    class ConditionStringBuilderTests_Dao
    {
        private IConditionStringBuilder _csb;

        [SetUp]
        public void MyTestInitialize()
        {
            _csb = new ConditionStringBuilder();
            _csb.SqlConverter = new DaoSqlConverter();
        }

        [TearDown]
        public void MyTestCleanup()
        {
            _csb = null;
        }

        [Test]
        public void GenerateSqlString_WithoutConditions_ReturnsEmpty()
        {
            var actual = _csb.ToString();
            Assert.AreEqual(string.Empty, actual);
        }

        [Test]
        [TestCase("Field1", FieldDataType.Numeric, RelationalOperators.Equal, 123, "Field1 = 123")]
        [TestCase("Field1", FieldDataType.Text, RelationalOperators.Like, "abc", "Field1 Like 'abc'")]
        [TestCase("Field1", FieldDataType.Text, RelationalOperators.Like | RelationalOperators.AddWildcardSuffix, "abc", "Field1 Like 'abc*'")] 
        [TestCase("Field1", FieldDataType.Numeric, RelationalOperators.Equal, 123, "Field1 = 123")]
        [TestCase("Field1", FieldDataType.Boolean, RelationalOperators.Equal, true, "Field1 = True")]
        public void GenerateSqlString_FieldAndValue(object fieldName, FieldDataType dataType, RelationalOperators relationalOperator, object value, string expected)
        {
            _csb.Add(fieldName, dataType, relationalOperator, value);

            var actual = _csb.ToString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [TestCase("Field1", RelationalOperators.Equal, "2015-12-24", "Field1 = #2015-12-24#")]
        [TestCase("Field1", RelationalOperators.Equal, "2015-12-24 12:30:50", "Field1 = #2015-12-24 12:30:50#")]
        public void GenerateSqlStringFieldAndValue_Date(object fieldName, RelationalOperators relationalOperator, string value, string expected)
        {
            System.DateTime date = System.DateTime.Parse(value);

            _csb.Add(fieldName, FieldDataType.DateTime, relationalOperator, date);

            var actual = _csb.ToString();
            Assert.AreEqual(expected, actual);
        }

        [TestCase("Field1", FieldDataType.Numeric, RelationalOperators.Between, 1, 5, "(Field1 Between 1 And 5)")]
        [TestCase("Field1", FieldDataType.Text, RelationalOperators.Between, "a", "x", "(Field1 Between 'a' And 'x')")]
        [TestCase("Field1", FieldDataType.Text, RelationalOperators.Between | RelationalOperators.AddWildcardSuffix, "a", "x", "(Field1 >= 'a' And Field1 < 'y')")]        
        public void GenerateSqlStringFieldAndValue_Between(object fieldName, FieldDataType fieldDataType, RelationalOperators relationalOperator, object value1, object value2, string expected)
        {
            object[] values = new[] { value1, value2 };

            _csb.Add(fieldName, fieldDataType, relationalOperator, values);

            var actual = _csb.ToString();
            Assert.AreEqual(expected, actual);
        }

        [TestCase("Field1", RelationalOperators.Between, "2015-01-01", "2015-12-24 12:30:50", "(Field1 Between #2015-01-01# And #2015-12-24 12:30:50#)")]
        [TestCase("Field1", RelationalOperators.Between | RelationalOperators.AddWildcardSuffix, "2015-01-01", "2015-12-24 12:30:50", "(Field1 >= #2015-01-01# And Field1 < #2015-12-25#)")]
        public void GenerateSqlStringFieldAndValue_Date_Between(object fieldName, RelationalOperators relationalOperator, string value1, string value2, string expected)
        {
            System.DateTime date1 = System.DateTime.Parse(value1);
            System.DateTime date2 = System.DateTime.Parse(value2);

            var dates = new System.DateTime[] { date1, date2};

            _csb.Add(fieldName, FieldDataType.DateTime, relationalOperator, dates);

            var actual = _csb.ToString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [TestCase(FieldDataType.Numeric, RelationalOperators.Equal)]
        [TestCase(FieldDataType.Text, RelationalOperators.Equal)]
        [TestCase(FieldDataType.DateTime, RelationalOperators.Equal)]
        public void GenerateSqlString_DbNullValue(FieldDataType dataType, RelationalOperators relationalOperator)
        {
            _csb.Add("Field1", dataType, relationalOperator, System.DBNull.Value);
            _csb.Add("Field2", dataType, relationalOperator, System.DBNull.Value);

            var actual = _csb.ToString();
            Assert.AreEqual(string.Empty, actual);
        }

        [Test]
        [TestCase(FieldDataType.Numeric, RelationalOperators.Equal)]
        [TestCase(FieldDataType.Text, RelationalOperators.Equal)]
        [TestCase(FieldDataType.DateTime, RelationalOperators.Equal)]
        public void GenerateSqlString_ConcatDbNullValue(FieldDataType dataType, RelationalOperators relationalOperator)
        {
            _csb.Add("Field1", FieldDataType.Text, RelationalOperators.Between, new object[]{ System.DBNull.Value, "xyz"});
            _csb.Add("Field2", FieldDataType.Numeric, RelationalOperators.Between, new object[] {133.45, System.DBNull.Value});
            _csb.Add("Field3", dataType, relationalOperator, System.DBNull.Value);

            var actual = _csb.ToString();
            Assert.AreEqual("Field1 <= 'xyz' And Field2 >= 133.45", actual);
        }

        [Test]
        [TestCase(FieldDataType.Numeric)]
        [TestCase(FieldDataType.Text)]
        [TestCase(FieldDataType.DateTime)]
        public void GenerateSqlString_ConcatBetweenDbNullValue(FieldDataType dataType)
        {
            _csb.Add("Field1", FieldDataType.Text, RelationalOperators.Between, new object[] { System.DBNull.Value, "xyz" });
            _csb.Add("Field2", FieldDataType.Numeric, RelationalOperators.Between, new object[] { 133.45, System.DBNull.Value });
            _csb.Add("Field3", dataType, RelationalOperators.Between, new object[]{System.DBNull.Value, System.DBNull.Value});

            var actual = _csb.ToString();
            Assert.AreEqual("Field1 <= 'xyz' And Field2 >= 133.45", actual);
        }

        [Test]
        [TestCase(FieldDataType.Numeric, 1, 3, "(Field1 Between 1 And 3)")]
        [TestCase(FieldDataType.Text, "a", "c", "(Field1 Between 'a' And 'c')")]
        public void GenerateBetweenSqlString_Values(FieldDataType dataType, object value1, object value2, string expected)
        {
            var betweenValues = new object[] { value1, value2 };

            _csb.Add("Field1", dataType, RelationalOperators.Between, betweenValues);

            var actual = _csb.ToString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [TestCase(FieldDataType.Numeric, 1, 3, "(Field1 Between 1 And 3)")]
        [TestCase(FieldDataType.Text, "a", "c", "(Field1 Between 'a' And 'c')")]
        public void GenerateBetweenSqlString_AddBetweenCondition(FieldDataType dataType, object value1, object value2, string expected)
        {
            _csb.AddBetweenCondition("Field1", dataType, value1, value2);

            var actual = _csb.ToString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [TestCase(FieldDataType.Numeric)]
        [TestCase(FieldDataType.Text)]
        [TestCase(FieldDataType.DateTime)]
        public void GenerateBetweenSqlString_DbNullValue(FieldDataType dataType)
        {
            var betweenValues = new object[] { System.DBNull.Value, System.DBNull.Value };

            _csb.Add("Field1", dataType, RelationalOperators.Between, betweenValues);
            _csb.Add("Field2", dataType, RelationalOperators.Between, betweenValues);

            var actual = _csb.ToString();
            Assert.AreEqual(string.Empty, actual);
        }

        [Test]
        [TestCase("Field1", FieldDataType.Numeric, RelationalOperators.Equal, 123, 123, "")]
        [TestCase("Field1", FieldDataType.Numeric, RelationalOperators.Equal, null, 123, "Field1 Is Null")]
        [TestCase("Field1", FieldDataType.Numeric, RelationalOperators.Equal | RelationalOperators.GreaterThan, null, 123, "(Field1 Is Null Or Field1 > Null)")]
        [TestCase("Field1", FieldDataType.Text, RelationalOperators.Equal, "abc", "abc", "")]
        [TestCase("Field1", FieldDataType.Text, RelationalOperators.Equal, "", null, "Field1 = ''")]
        [TestCase("Field1", FieldDataType.Text, RelationalOperators.Equal, "", "", "")]
        [TestCase("Field1", FieldDataType.Text, RelationalOperators.Equal, null, "", "Field1 Is Null")]
        public void GenerateSqlString_WithIgnoreValue(string fieldName, FieldDataType dataType, RelationalOperators relationalOperator, object value, object ignoreValue, string expected)
        {
            _csb.Add(fieldName, dataType, relationalOperator, value, ignoreValue);

            var actual = _csb.ToString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [TestCase("Field1",RelationalOperators.Equal, "2015-12-24", "2015-12-24", "")]
        public void GenerateSqlString_WithIgnoreValue_Date(string fieldName, RelationalOperators relationalOperator, string value, string ignoreValue, string expected)
        {
            System.DateTime date = System.DateTime.Parse(value);
            System.DateTime ignoreDate = System.DateTime.Parse(ignoreValue);

            _csb.Add(fieldName, FieldDataType.DateTime, relationalOperator, date, ignoreDate);

            var actual = _csb.ToString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void GenerateSqlString_InStatement_Numeric()
        {
            const string expected = "F In (1,3,5)";

            var values = new int[] { 1, 3, 5 };
            _csb.Add("F", FieldDataType.Numeric, RelationalOperators.In, values);

            var actual = _csb.ToString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void GenerateSqlString_InStatement_NullValuesInArray()
        {
            const string expected = "";

            var values = new[] { System.DBNull.Value, System.DBNull.Value, System.DBNull.Value };
            _csb.Add("F", FieldDataType.Numeric, RelationalOperators.In, values, System.DBNull.Value);

            var actual = _csb.ToString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void GenerateSqlString_EqualValueArray_Numeric()
        {
            const string expected = "(F = 1 Or F = 3 Or F = 5)";

            var values = new int[] { 1, 3, 5 };
            _csb.Add("F", FieldDataType.Numeric, RelationalOperators.Equal, values);

            var actual = _csb.ToString();
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void GenerateSqlString_EqualValueArray_TextWithWildcard()
        {
            const string expected = "(F Like 'a*' Or F Like 'c*' Or F Like 'e*')";

            var values = new string[] { "a", "c", "e" };
            _csb.Add("F", FieldDataType.Text, RelationalOperators.Like | RelationalOperators.AddWildcardSuffix, values);

            var actual = _csb.ToString();
            Assert.AreEqual(expected, actual);
        }


        [Test]
        public void GenerateSqlString_ConditionGroup()
        {
            const string expected = "F1 = 1 And (F2 = 2 Or F3 = 3)";

            var values = new string[] { "a", "c", "e" };
            _csb.Add("F1", FieldDataType.Numeric, RelationalOperators.Equal, 1);

            _csb.BeginGroup(LogicalOperator.Or)
                .Add("F2", FieldDataType.Numeric, RelationalOperators.Equal, 2)
                .Add("F3", FieldDataType.Numeric, RelationalOperators.Equal, 3);

            var actual = _csb.ToString();
            Assert.AreEqual(expected, actual);
        }
    }
}
