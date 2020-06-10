using NUnit.Framework;

namespace AccessCodeLib.Data.SqlTools.interop.Tests
{
    abstract class GeneratorTestBase<T>
    {
        protected T Generator { get; private set; }
        protected abstract T GetGenerator();
        
        [SetUp]
        public void MyTestInitialize()
        {
            Generator = GetGenerator();
        }

        [TearDown]
        public void MyTestCleanup()
        {
            Generator = default(T);
        }
    }
}
