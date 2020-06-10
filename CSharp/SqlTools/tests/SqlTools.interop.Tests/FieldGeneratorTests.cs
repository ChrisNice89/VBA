using System.Linq;
using NUnit.Framework;

namespace AccessCodeLib.Data.SqlTools.interop.Tests
{
    [TestFixture]
    public class FieldGeneratorTests
    {
        private FieldGenerator _fieldGenerator;

        [SetUp]
        public void Setup()
        {
            _fieldGenerator = new FieldGenerator();
        }

        [TearDown]
        public void TearDown()
        {
            _fieldGenerator = null;
        }

        [Test]
        public void FromStringArrayTest()
        {
            var fields = new[] { "F1", "F2", "F3" };
            var expected = new[] { new Field("F1"), new Field("F2"), new Field("F3") };
            var actual = (_fieldGenerator.FromArray(fields).Cast<IField>().ToArray());

            Assert.AreEqual(expected.Count(), actual.Count());
            Assert.AreEqual(expected[0].Name, actual[0].Name);
            Assert.AreEqual(expected[1].Name, actual[1].Name);
            Assert.AreEqual(expected[2].Name, actual[2].Name);
        }

        [Test]
        public void FromFieldArrayTest()
        {
            var fields = new[] { new Field("F1"), new Field("F2"), new Field("F3") };
            var expected = new[] { new Field("F1"), new Field("F2"), new Field("F3") };
            var actual = _fieldGenerator.FromArray(fields).Cast<IField>().ToArray();

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
            var actual = _fieldGenerator.FromString(fieldsString).Cast<IField>().ToArray();

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
            var actual = _fieldGenerator.FromString(fieldsString, ';').ToArray();

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

            var actual = _fieldGenerator.FromString(fieldsString, ";").Cast<IField>().ToArray();

            Assert.AreEqual(expected.Count(), actual.Count());
            Assert.AreEqual(expected[0].Name, actual[0].Name);
            Assert.AreEqual(expected[1].Name, actual[1].Name);
        }
    }
}
