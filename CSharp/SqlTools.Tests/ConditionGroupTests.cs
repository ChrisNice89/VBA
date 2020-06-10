using AccessCodeLib.Data.SqlTools.Sql;
using NUnit.Framework;
using AccessCodeLib.Data.Common.Sql;

namespace AccessCodeLib.Data.SqlTools.Tests
{
    [TestFixture]
    class ConditionGroupTests
    {
        private ConditionGroup _conditionGroup;

        [SetUp]
        public void SetUp()
        {
            _conditionGroup = new ConditionGroup();
        }

        [TearDown]
        public void TearDown()
        {
            _conditionGroup = null;
        }

        [Test]
        public void Add_1FieldCondition_CheckValue()
        {
            var expected = new FieldCondition(new Field("F1"), RelationalOperators.Equal, 123);

            _conditionGroup.Add(new Field("F1"), RelationalOperators.Equal, 123);

            var actual = (IFieldCondition)_conditionGroup.Conditions[0];
            Assert.AreEqual(expected.Value, actual.Value);
        }

        [Test]
        public void Add_2FieldConditions_CheckValueItem2()
        {
            var expected = new FieldCondition(new Field("F2"), RelationalOperators.GreaterThan, 0);

            _conditionGroup.Add(new Field("F1"), RelationalOperators.Equal, 123)
                           .Add(new Field("F2"), RelationalOperators.GreaterThan, 0);

            var actual = (IFieldCondition)(_conditionGroup.Conditions[1]);
            Assert.AreEqual(expected.Operator, actual.Operator);
        }

        [Test]
        public void Add_1ConditionGroup_CheckValue()
        {
            var expected = new FieldCondition(new Field("F1"), RelationalOperators.Equal, 123);
            var groupToAdd = new ConditionGroup(new [] {expected});

            _conditionGroup.Add(groupToAdd);

            var actual = (IFieldCondition)((IConditionGroup)_conditionGroup.Conditions[0]).Conditions[0];
            Assert.AreEqual(expected.Value, actual.Value);
        }

        [Test]
        public void Add_2FieldsInConditionGroup_CheckValue()
        {
            var expected = new FieldCondition(new Field("F1"), RelationalOperators.Equal, new Field("F2"));

            _conditionGroup.Add(new Field("F1"), RelationalOperators.Equal, new Field("F2"));

            var actual = ((IFieldCondition)_conditionGroup.Conditions[0]);
            Assert.AreEqual(expected.Value.GetType(), actual.Value.GetType());

        }
    }
}