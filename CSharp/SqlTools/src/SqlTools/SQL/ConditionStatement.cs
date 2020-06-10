using AccessCodeLib.Data.Common.Sql;

namespace AccessCodeLib.Data.SqlTools.Sql
{
    public abstract class ConditionStatement : ConditionGroup, IStatement
    {
        protected ConditionStatement()
        {
        }

        protected ConditionStatement(ICondition condition)
        {
            Add(condition);
        }

        protected ConditionStatement(string condition)
        {
            Add(new ConditionString(condition));
        }

        protected ConditionStatement(IField field, RelationalOperators relationalOperator, object value)
        {
            Add(field, relationalOperator, value);
        }

        public abstract string Key { get; }
    }
}