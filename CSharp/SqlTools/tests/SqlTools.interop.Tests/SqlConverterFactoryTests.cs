using AccessCodeLib.Data.SqlTools.Converter.Mssql;
using NUnit.Framework;

namespace AccessCodeLib.Data.SqlTools.interop.Tests
{
// ReSharper disable UnusedMember.Global
    class SqlConverterFactoryTests
// ReSharper restore UnusedMember.Global
    {
        private SqlConverterFactory _factory;

        [SetUp]
        public void MyTestInitialize()
        {
            _factory = new SqlConverterFactory();
        }

        [TearDown]
        public void MyTestCleanup()
        {
            _factory = null;
        }

        [Test]
        public void Ansi92SqlConverter_CheckType()
        {
            var actual = _factory.Ansi92SqlConverter();

            Assert.That(actual, Is.InstanceOf(typeof(ISqlConverter)));
            Assert.That(actual, Is.InstanceOf(typeof(Converter.Common.Ansi92.SqlConverter)));
            Assert.That(actual, Is.InstanceOf(typeof(Ansi92SqlConverter)));
        }

        [Test]
        public void DaoSqlConverter_CheckType()
        {
            var actual = _factory.DaoSqlConverter();

            Assert.That(actual, Is.InstanceOf(typeof(ISqlConverter)));
            Assert.That(actual, Is.InstanceOf(typeof(Converter.Jet.Dao.SqlConverter)));
            Assert.That(actual, Is.InstanceOf(typeof(DaoSqlConverter)));
        }

        [Test]
        public void JetAdodbSqlConverter_CheckType()
        {
            var actual = _factory.JetAdodbSqlConverter();

            Assert.That(actual, Is.InstanceOf(typeof(ISqlConverter)));
            Assert.That(actual, Is.InstanceOf(typeof(Converter.Jet.Oledb.SqlConverter)));
            Assert.That(actual, Is.InstanceOf(typeof(JetAdodbSqlConverter)));
        }

        [Test]
        public void TsqlSqlConverter_CheckType()
        {
            var actual = _factory.TsqlSqlConverter();

            Assert.That(actual, Is.InstanceOf(typeof(ISqlConverter)));
            Assert.That(actual, Is.InstanceOf(typeof(SqlConverter)));
            Assert.That(actual, Is.InstanceOf(typeof(TsqlSqlConverter)));
        }
    }
}
