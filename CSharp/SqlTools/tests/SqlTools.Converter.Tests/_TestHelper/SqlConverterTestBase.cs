using AccessCodeLib.Data.Common.Sql.Converter;
using NUnit.Framework;

namespace AccessCodeLib.Data.SqlTools.Converter.Tests
{
    abstract class SqlConverterTestBase
    {
        protected ISqlGenerator Generator { get; private set; }
        private ISqlConverter Converter { get; set; }
        protected abstract ISqlConverter GetConverter();

        private string GenerateSqlString(ISqlGenerator generator)
        {
            return Converter.GenerateSqlString(generator.SqlStatement);
        }

        protected string GenerateSqlString()
        {
            return GenerateSqlString(Generator);
        }

        [SetUp]
        public void MyTestInitialize()
        {
            Converter = GetConverter();
            Generator = new SqlGenerator();
        }

        [TearDown]
        public void MyTestCleanup()
        {
            Converter = null;
        }

        // Allgemeine Tests, die mit jedem Converter funktionieren müssen

        [Test]
        public void CreateGenerator_InitWithConverter()
        {
            var g = new SqlGenerator(Converter);
            Assert.That(g.Converter, Is.SameAs(Converter));
        }

        [Test]
        public void CreateGenerator_AppendConverter()
        {
            var g = new SqlGenerator {Converter = Converter};
            Assert.That(g.Converter, Is.SameAs(Converter));
        }

        [Test]
        public void GenerateSqlString_StatementIsNull_ReturnsNull()
        {
            var sqlString = Converter.GenerateSqlString(null);
            Assert.That(sqlString, Is.Null);
        }

        [Test]
        public void GenerateSqlString_StatementIsNotNull_ReturnsEmptyString()
        {
            var statement = Generator.SqlStatement;
            var sqlString = Converter.GenerateSqlString(statement);
            Assert.That(sqlString, Is.Empty);
        }

        [Test]
        public void GenerateSqlString_FromStatementWithTableName_ReturnsNotEmptyString()
        {
            var statement = Generator.From("Table").SqlStatement;
            var sqlString = Converter.GenerateSqlString(statement);
            Assert.That(sqlString, Is.Not.Empty);
        }
    }
}
