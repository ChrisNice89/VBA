using AccessCodeLib.Data.Common.Sql;

namespace AccessCodeLib.Data.SqlTools
{
    public interface IConditionGenerator : IConditionGroup
    {
        IConditionGroup BeginGroup(LogicalOperator concatOperator = LogicalOperator.And);
    }
}