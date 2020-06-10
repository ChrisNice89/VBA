using AccessCodeLib.Data.SqlTools.Sql;
using NUnit.Framework;
using AccessCodeLib.Data.Common.Sql;

namespace AccessCodeLib.data.SqlTools.Tests.Sql
{
    [TestFixture]
    class BetweenValueTests
    {
        [Test]
        public void IntValue_CheckValues()
        {
            var betweenValue = new BetweenValue(1, 5);
            Assert.AreEqual(1, ((INumericValue<int>)betweenValue.FirstValue).Value);
            Assert.AreEqual(5, ((INumericValue<int>)betweenValue.SecondValue).Value);
        }
    }
}
