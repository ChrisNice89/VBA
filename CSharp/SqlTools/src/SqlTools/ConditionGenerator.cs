using AccessCodeLib.Data.Common.Sql;
using AccessCodeLib.Data.SqlTools.Sql;

namespace AccessCodeLib.Data.SqlTools
{
    public class ConditionGenerator : ConditionGroup, IConditionGenerator
    {
        public IConditionGroup BeginGroup(LogicalOperator concatOperator = LogicalOperator.And)
        {
            var group = new ConditionGroup(concatOperator);
            Add(group);
            return group;
        }
    }
}