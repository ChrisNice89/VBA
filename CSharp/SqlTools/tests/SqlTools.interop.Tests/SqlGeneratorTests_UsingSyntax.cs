using AccessCodeLib.Data.SqlTools.Sql;
using NUnit.Framework;

namespace AccessCodeLib.Data.SqlTools.interop.Tests
{
    [TestFixture]
    class SqlGeneratorTests_UsingSyntax : GeneratorTestBase<ISqlGenerator>
    {
        protected override ISqlGenerator GetGenerator()
        {
            return new SqlGenerator(new DaoSqlConverter());
        }

        #region Select

        [Test]
        public void Select_FieldNames()
        {
            const string expected = "Select F1, F2, F3";

            Generator.Select("F1", "F2", "F3");

            Assert.AreEqual(expected, Generator.ToString());
        }

        [Test]
        public void Select_FieldNameArray()
        {
            const string expected = "Select F1, F2, F3";

            var fields = new[] {"F1", "F2", "F3"};
            Generator.Select(fields);

            Assert.AreEqual(expected, Generator.ToString());
        }

        [Test]
        public void Select_FieldArray()
        {
            const string expected = "Select F1, F2, F3";

            var fieldGenerator = new FieldGenerator();
            var fields = fieldGenerator.FromArray(new[] {"F1", "F2", "F3"});
            Generator.Select(fields);

            Assert.AreEqual(expected, Generator.ToString());
        }

        [Test]
        public void Select_FieldArray_WithSource()
        {
            const string expected = "Select Tab1.F1, Tab1.F2, Tab1.F3";

            var fieldGenerator = new FieldGenerator();
            var fields = fieldGenerator.FromArray(new[] {
                                                            new Field("F1", "Tab1"), 
                                                            new Field("F2", "Tab1"), 
                                                            new Field("F3", "Tab1")
                                                        });
            Generator.Select(fields);

            Assert.AreEqual(expected, Generator.ToString());
        }

        [Test]
        public void SelectAll()
        {
            const string expected = "Select *";

            Generator.SelectAll();

            Assert.AreEqual(expected, Generator.ToString());
        }

        [Test]
        public void SelectField_WithSourceAndAlias()
        {
            const string expected = "Select Tab1.F1 As T1F1, Tab2.F1 As T2F1";

            Generator.SelectField("F1", "Tab1", "T1F1")
                     .SelectField("F1", "Tab2", "T2F1");

            Assert.AreEqual(expected, Generator.ToString());
        }

        #endregion

        #region From

        [Test]
        public void From_SourceNames()
        {
            const string expected = "From Tab1, Tab2";

            Generator.From("Tab1").From("Tab2");

            Assert.AreEqual(expected, Generator.ToString());
        }

        [Test]
        public void From_SourceNames_InnerJoin()
        {
            const string expected = "From Tab1 Inner Join Tab2 On Tab1.F1 = Tab2.F2";

            Generator.From("Tab1").InnerJoin("Tab1", "Tab2", "F1", "F2");

            Assert.AreEqual(expected, Generator.ToString());
        }

        [Test]
        public void From_SourceNames_LeftJoin()
        {
            const string expected = "From Tab1 Left Join Tab2 On Tab1.F1 = Tab2.F2";

            Generator.From("Tab1").LeftJoin("Tab1", "Tab2", "F1", "F2");

            Assert.AreEqual(expected, Generator.ToString());
        }

        [Test]
        public void From_SourceNames_RightJoin()
        {
            const string expected = "From Tab1 Right Join Tab2 On Tab1.F1 = Tab2.F2";

            Generator.From("Tab1").RightJoin("Tab1", "Tab2", "F1", "F2");

            Assert.AreEqual(expected, Generator.ToString());
        }

        [Test]
        public void From_SourceNames_2InnerJoin()
        {
            const string expected = "From (Tab1 Inner Join Tab2 On Tab1.F1 = Tab2.F2) Inner Join Tab3 On Tab1.F1 >= Tab3.F2";

            Generator.From("Tab1")
                     .InnerJoin("Tab1", "Tab2", "F1", "F2")
                     .InnerJoin("Tab1", "Tab3", "F1", "F2", RelationalOperators.Equal | RelationalOperators.GreaterThan);

            Assert.AreEqual(expected, Generator.ToString());
        }

        [Test]
        public void From_SourceNames_CrossAndInnerJoin()
        {
            const string expected = "From TabA, TabB Inner Join Tab2 On TabB.F1 = Tab2.F2";

            Generator.From("TabA").From("TabB").InnerJoin("TabB", "Tab2", "F1", "F2");

            Assert.AreEqual(expected, Generator.ToString());
        }

        [Test]
        public void From_SourceNames_CrossAndInnerJoin_WithAlias()
        {
            const string expected = "From TabA As A, TabB As B Inner Join TabC As C On B.F1 = C.F2";

            var sourceA = new SourceAlias(new NamedSource("TabA"), "A");
            var sourceB = new SourceAlias(new NamedSource("TabB"), "B");
            var sourceC = new SourceAlias(new NamedSource("TabC"), "C");

            Generator.From(sourceA).From(sourceB).InnerJoin(sourceB, sourceC, "F1", "F2");

            Assert.AreEqual(expected, Generator.ToString());
        }

        [Test]
        public void From_SourceNames_2InnerJoin_WithAlias()
        {
            const string expected = "From (TabA As A Inner Join TabB As B On A.F1 = B.F2) Inner Join TabC As C On A.F1 = C.F2";

            var sourceA = new SourceAlias(new NamedSource("TabA"), "A");
            var sourceB = new SourceAlias(new NamedSource("TabB"), "B");
            var sourceC = new SourceAlias(new NamedSource("TabC"), "C");

            Generator.From(sourceA).InnerJoin(sourceA, sourceB, "F1", "F2");
            Generator.InnerJoin(sourceA, sourceC, "F1", "F2");

            Assert.AreEqual(expected, Generator.ToString());
        }

        [Test]
        public void From_SourceNames_3InnerJoin()
        {
            const string expected = "From ((Tab1 Inner Join Tab2 On Tab1.F1 = Tab2.F2) Inner Join Tab3 On Tab1.F1 = Tab3.F2) Inner Join Tab4 On Tab2.F1 = Tab4.F2";

            Generator.From("Tab1").InnerJoin("Tab1", "Tab2", "F1", "F2");
            Generator.InnerJoin("Tab1", "Tab3", "F1", "F2");
            Generator.InnerJoin("Tab2", "Tab4", "F1", "F2");

            Assert.AreEqual(expected, Generator.ToString());
        }

        [Test]
        public void From_SourceNames_InnerJoin_complex()
        {
            const string expected = "From (Tab1 Inner Join Tab2 On Tab1.F1 > Tab2.F2 And Tab2.F3 >= 5) Inner Join Tab3 On Tab1.F1 = Tab3.F2";

            Generator.From("Tab1").InnerJoin("Tab1", "Tab2", "F1", "F2", RelationalOperators.GreaterThan, 
                                                                5, "F3", RelationalOperators.LessThan | RelationalOperators.Equal);
            Generator.InnerJoin("Tab1", "Tab3", "F1", "F2");
            
            Assert.AreEqual(expected, Generator.ToString());
        }

        #endregion

        #region Where

        [Test]
        public void Where_FieldRelationalOperatorValue()
        {
            const string expected = "Where F1 = 5 And F2 > 'a'";

            Generator.Where("F1", RelationalOperators.Equal, 5)
                     .Where("F2",RelationalOperators.GreaterThan, "a");

            Assert.AreEqual(expected, Generator.ToString());
        }

        [Test]
        public void Where_FieldRelationalOperatorField()
        {
            const string expected = "Where F1 = F2";

            var fieldGenerator = new FieldGenerator();
            var field = fieldGenerator.Field("F2");
            Generator.Where("F1", RelationalOperators.Equal, field);

            Assert.AreEqual(expected, Generator.ToString());
        }

        [Test]
        public void Where_ConditionParam()
        {
            const string expected = "Where (F1 = 5 And F2 > 'a')";

            var conditionGenerator = new ConditionGenerator();
            var condition = conditionGenerator.Add("F1", FieldDataType.Numeric, RelationalOperators.Equal, 5)
                                              .Add("F2", FieldDataType.Text, RelationalOperators.GreaterThan, "a");
            
            Generator.WhereCondition(condition);

            Assert.AreEqual(expected, Generator.ToString());
        }

        [Test]
        public void Where_WhereString()
        {
            const string expected = "Where (F1 = 5 And F2 > 'a')";

            Generator.WhereString("F1 = 5 And F2 > 'a'");

            Assert.AreEqual(expected, Generator.ToString());
        }

        [Test]
        public void WhereBetweenString()
        {
            const string expected = "Where (F1 Between 1 And 9)";

            Generator.WhereBetween("F1", 1, 9);

            Assert.AreEqual(expected, Generator.ToString());
        }

        [Test]
        public void WhereBetween_IntWithNull_ToString()
        {
            const string expected = "Where F1 >= 1";

            Generator.WhereBetween("F1", 1, System.DBNull.Value);

            Assert.AreEqual(expected, Generator.ToString());
        }

        [Test]
        public void WhereBetween_StringWithNull_ToString()
        {
            const string expected = "Where F1 >= '1'";

            Generator.WhereBetween("F1", "1", System.DBNull.Value);

            Assert.AreEqual(expected, Generator.ToString());
        }

        #endregion

        #region GroupBy

        //
        // analog Select (nur ohne Alias)
        //

        [Test]
        public void GroupBy_FieldNames()
        {
            const string expected = "Group By F1, F2, F3";

            Generator.GroupBy("F1", "F2", "F3");

            Assert.AreEqual(expected, Generator.ToString());
        }

        #endregion

        #region Having

        //
        // analog Where
        //

        [Test]
        public void Having_FieldRelationalOperatorValue()
        {
            const string expected = "Having F1 = 5 And F2 > 'a'";

            Generator.Having("F1", RelationalOperators.Equal, 5)
                     .Having("F2", RelationalOperators.GreaterThan, "a");

            Assert.AreEqual(expected, Generator.ToString());
        }

        [Test]
        public void Having_ConditionParam()
        {
            const string expected = "Having (F1 = 5 And F2 > 'a')";

            var conditionGenerator = new ConditionGenerator();
            var condition = conditionGenerator.Add("F1", FieldDataType.Numeric, RelationalOperators.Equal, 5)
                                              .Add("F2", FieldDataType.Text, RelationalOperators.GreaterThan, "a");

            Generator.HavingCondition(condition);

            Assert.AreEqual(expected, Generator.ToString());
        }

        [Test]
        public void Having_HavingString()
        {
            const string expected = "Having (F1 = 5 And F2 > 'a')";

            Generator.HavingString("F1 = 5 And F2 > 'a'");

            Assert.AreEqual(expected, Generator.ToString());
        }

        #endregion

        #region OrderBy

        //
        // analog Select (nur ohne Alias)
        //

        [Test]
        public void OrderBy_FieldNames()
        {
            const string expected = "Order By F1, F2, F3";

            Generator.OrderBy("F1", "F2", "F3");

            Assert.AreEqual(expected, Generator.ToString());
        }

        #endregion
    }
}
