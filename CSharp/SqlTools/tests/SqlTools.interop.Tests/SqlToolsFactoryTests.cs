using NUnit.Framework;

namespace AccessCodeLib.Data.SqlTools.interop.Tests
{
    [TestFixture]
    public class SqlToolsFactoryTests
    {
        private SqlToolsFactory _factory;

        [SetUp]
        public void MyTestInitialize()
        {
            _factory = new SqlToolsFactory();
        }

        [TearDown]
        public void MyTestCleanup()
        {
            _factory = null;
        }

        [Test]
        public void FieldGenerator_CheckType()
        {
            var actual = _factory.FieldGenerator();

            Assert.That(actual, Is.InstanceOf(typeof (IFieldGenerator)));

            Assert.That(actual, Is.InstanceOf(typeof(SqlTools.IFieldGenerator)));
        }

        [Test]
        public void ConditionGenerator_CheckType()
        {
            var actual = _factory.ConditionGenerator();

            Assert.That(actual, Is.InstanceOf(typeof(IConditionGenerator)));
            Assert.That(actual, Is.InstanceOf(typeof(IConditionGroup)));

            Assert.That(actual, Is.InstanceOf(typeof(SqlTools.IConditionGenerator)));
        }

        [Test]
        public void SqlConverters_CheckType()
        {
            var actual = _factory.SqlConverters;

            Assert.That(actual, Is.InstanceOf(typeof(ISqlConverterFactory)));
        }

        [Test]
        public void SqlGenerator_CheckType()
        {
            var actual = _factory.SqlGenerator();

            Assert.That(actual, Is.InstanceOf(typeof(ISqlGenerator)));
            Assert.That(actual, Is.InstanceOf(typeof(SqlTools.ISqlGenerator)));
        }
    }
}
