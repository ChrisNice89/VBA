using AccessCodeLib.Data.Common.Sql;
using NUnit.Framework;

namespace AccessCodeLib.Data.SqlTools.Tests
{
    [TestFixture]
    class RelationalOperatorTests
    {
        [Test]
        public void EqualAndLessThan_CheckEqual()
        {
            const RelationalOperators x = RelationalOperators.Equal | RelationalOperators.LessThan;
            Assert.True ((x & RelationalOperators.Equal) == RelationalOperators.Equal);
        }
    }
}
