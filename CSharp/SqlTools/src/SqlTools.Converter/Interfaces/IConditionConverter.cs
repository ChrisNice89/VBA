using System.Collections.Generic;
using AccessCodeLib.Data.Common.Sql;

namespace AccessCodeLib.Data.SqlTools.Converter
{
    public interface IConditionConverter
    {
        string GenerateSqlString(IEnumerable<IConditionGroup> conditionGroups, LogicalOperator topLevelConcatOperator = LogicalOperator.And);
        string GenerateSqlString(IConditionGroup conditionGroup);
        string GenerateSqlString(string conditionString);
    }
}