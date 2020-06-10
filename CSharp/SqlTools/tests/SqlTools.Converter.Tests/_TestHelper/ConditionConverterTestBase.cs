using System.Collections.Generic;
using AccessCodeLib.Data.Common.Sql;
using AccessCodeLib.Data.SqlTools.Sql;
using NUnit.Framework;

namespace AccessCodeLib.Data.SqlTools.Converter.Tests
{
    abstract class ConditionConverterTestBase
    {
        private IConditionConverter Converter { get; set; }
        protected abstract IConditionConverter GetConverter();

        protected ConditionGroup Conditions;
        private List<IConditionGroup> ConditionGroups;

        protected string GenerateSqlString()
        {
            return Converter.GenerateSqlString(ConditionGroups);
        }
        
        [SetUp]
        public void Setup()
        {
            Converter = GetConverter();
            Conditions = new ConditionGroup();
            ConditionGroups = new List<IConditionGroup> { Conditions };
        }

        [TearDown]
        public void TearDown()
        {
            Converter = null;
            Conditions = null;
        }
    }
}
