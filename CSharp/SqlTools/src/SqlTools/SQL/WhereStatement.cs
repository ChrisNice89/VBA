using AccessCodeLib.Data.Common.Sql;

namespace AccessCodeLib.Data.SqlTools.Sql
{
    public class WhereStatement : ConditionStatement, IWhereStatement
    {
        public WhereStatement()
        {
        }

        public WhereStatement(ICondition condition) : base(condition)
        {
        }

        public WhereStatement(IField field, RelationalOperators relationalOperator, object value)
            : base(field, relationalOperator, value)
        {
        }

        public WhereStatement(string condition)
            : base(condition)
        {
        }

        public override string Key
        {
            get { return StatementKeys.Where; }
        }
    }
}