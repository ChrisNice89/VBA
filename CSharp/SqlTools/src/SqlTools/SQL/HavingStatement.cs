using AccessCodeLib.Data.Common.Sql;

namespace AccessCodeLib.Data.SqlTools.Sql
{
    /*
    public class HavingStringStatement : ConditionStringStatement, IHavingStringStatement
    {
        public HavingStringStatement()
        {
        }

        public HavingStringStatement(string condition)
            : base(condition)
        {
        }

        public override string Key
        {
            get { return StatementKeys.Having; }
        }
    }
    */
    public class HavingStatement : ConditionStatement, IHavingStatement
    {
        public HavingStatement()
        {
        }

        public HavingStatement(ICondition condition)
            : base(condition)
        {
        }

        public HavingStatement(IField field, RelationalOperators relationalOperator, object value)
            : base(field, relationalOperator, value)
        {
        }

        public HavingStatement(string condition)
            : base(condition)
        {
        }

        public override string Key
        {
            get { return StatementKeys.Having; }
        }
    }
}