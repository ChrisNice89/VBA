using AccessCodeLib.Data.Common.Sql;
using AccessCodeLib.Data.Common.Sql.Converter;
using AccessCodeLib.Data.SqlTools.Sql;
using NUnit.Framework;

namespace AccessCodeLib.Data.SqlTools.Converter.Tests.Jet.Adodb
{
    [TestFixture]
    class SqlConverterTests : Jet.SqlConverterTests
    {
        protected override ISqlConverter GetConverter()
        {
            return new Converter.Jet.Oledb.SqlConverter();
        }

        [Test]
        public void Where_LikeWithWildcard_Ansi92Wildcard()
        {
            const string expected = "Where F1 Like 'a%'";
            Generator.Where(new Field("F1"), RelationalOperators.Like, "a*");
            var actual = GenerateSqlString();
            Assert.AreEqual(expected, actual);
        }
    }
}
