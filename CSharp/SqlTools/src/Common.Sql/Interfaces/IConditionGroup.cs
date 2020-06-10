using System.Collections.Generic;

namespace AccessCodeLib.Data.Common.Sql
{
    public interface IConditionGroup : ICondition
    {
// ReSharper disable UnusedMethodReturnValue.Global
        LogicalOperator ConcatOperator { get; set; }
        IList<ICondition> Conditions { get; }

        IConditionGroup Add(ICondition condition);
        IConditionGroup Add(IField field, RelationalOperators relationalOperator, object value, object IgnoreValue = null);
        IConditionGroup Add(IField field, RelationalOperators relationalOperator, IField field2);
        IConditionGroup Add(IConditionGroup group);
        IConditionGroup Add(IEnumerable<ICondition> conditions);
// ReSharper restore UnusedMethodReturnValue.Global
    }
}