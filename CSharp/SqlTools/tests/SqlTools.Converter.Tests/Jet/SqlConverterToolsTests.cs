using System;
using NUnit.Framework;

namespace AccessCodeLib.Data.SqlTools.Converter.Tests.Jet
{
    [TestFixture]
    //
    // Allgemeine Tests, die für DAO und ADODB gelten
    //
    public class SqlConverterToolsTests
    {
        [Test]
        [TestCase("abc", "abc")]
        [TestCase("ab c ", "[ab c ]")]
        [TestCase("ab-c", "[ab-c]")]
        [TestCase("ab+c", "[ab+c]")]
        [TestCase("ab*c", "[ab*c]")]
        [TestCase("ab=c", "[ab=c]")]
        [TestCase("ab/c", "[ab/c]")]
        [TestCase(@"ab\c", @"[ab\c]")]
        public void CheckedItemNameString(string name, string expected)
        {
            var actual = Converter.Jet.SqlConverterTools.CheckedItemNameString(name);
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void DateString_Date()
        {
            const string expected = "#2000-01-03#";

            var actual = Converter.Jet.SqlConverterTools.DateString(new DateTime(2000, 1, 3));
            Assert.AreEqual(expected, actual);
        }

        [Test]
        public void DateString_DateAndTime()
        {
            const string expected = "#2000-01-03 04:05:06#";

            var actual = Converter.Jet.SqlConverterTools.DateString(new DateTime(2000, 1, 3, 4, 5, 6));
            Assert.AreEqual(expected, actual);
        }
    }
}
