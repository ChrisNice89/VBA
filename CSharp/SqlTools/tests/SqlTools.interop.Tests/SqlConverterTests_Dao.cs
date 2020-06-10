using NUnit.Framework;

namespace AccessCodeLib.Data.SqlTools.interop.Tests
{
    [TestFixture]
    class SqlConverterTests_Dao
    {
        private ISqlConverter _converter;

        [SetUp]
        public void MyTestInitialize()
        {
            _converter = new DaoSqlConverter();
        }

        [TearDown]
        public void MyTestCleanup()
        {
            _converter = null;
        }

        [Test]
        public void GenerateSqlString_NullStatement_ReturnsNull()
        {
            var actual = _converter.GenerateSqlString(null);
            Assert.AreEqual(null, actual);
        }
    }
}
