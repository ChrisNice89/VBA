using System.Linq;
using AccessCodeLib.Data.SqlTools.Sql;
using NUnit.Framework;

namespace AccessCodeLib.Data.SqlTools.Tests
{
    [TestFixture]
    public class FieldGeneratorTests
    {
        private FieldGenerator FieldGenerator;

        [SetUp]
        public void Setup()
        {
            FieldGenerator = new FieldGenerator();
        }

        [TearDown]
        public void TearDown()
        {
            FieldGenerator = null;
        }

        [Test]
        public void FromArrayTest()
        {
            var fields = new[] { "F1", "F2", "F3" };
            var expected = new[] { new Field("F1"), new Field("F2"), new Field("F3") };
            var actual = FieldGenerator.FromArray(fields).ToArray();

            Assert.AreEqual(expected.Count(), actual.Count());
            Assert.AreEqual(expected[0].Name, actual[0].Name);
            Assert.AreEqual(expected[1].Name, actual[1].Name);
            Assert.AreEqual(expected[2].Name, actual[2].Name);
        }

        [Test]
        public void FromStringTest()
        {
            const string fieldsString = "F1, F2,F3";
            var expected = new[] { new Field("F1"), new Field("F2"), new Field("F3") };
            var actual = FieldGenerator.FromString(fieldsString).ToArray();

            Assert.AreEqual(expected.Count(), actual.Count());
            Assert.AreEqual(expected[0].Name, actual[0].Name);
            Assert.AreEqual(expected[1].Name, actual[1].Name);
            Assert.AreEqual(expected[2].Name, actual[2].Name);
        }

        [Test]
        public void FromString_WithDelimiterTest()
        {
            const string fieldsString = "F1; F2; F3";
            var expected = new[] { new Field("F1"), new Field("F2"), new Field("F3") };
            var actual = FieldGenerator.FromString(fieldsString, ';').ToArray();

            Assert.AreEqual(expected.Count(), actual.Count());
            Assert.AreEqual(expected[0].Name, actual[0].Name);
            Assert.AreEqual(expected[1].Name, actual[1].Name);
            Assert.AreEqual(expected[2].Name, actual[2].Name);
        }

        [Test]
        public void FromString_WithSeveralPossibleDelimitersTest()
        {
            const string fieldsString = "F1, F2; F3";
            var expected = new[] { new Field("F1, F2"), new Field("F3") };
            var actual = FieldGenerator.FromString(fieldsString, ';').ToArray();

            Assert.AreEqual(expected.Count(), actual.Count());
            Assert.AreEqual(expected[0].Name, actual[0].Name);
            Assert.AreEqual(expected[1].Name, actual[1].Name);
        }
    }
}
