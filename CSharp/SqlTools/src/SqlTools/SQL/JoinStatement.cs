using AccessCodeLib.Data.Common.Sql;

namespace AccessCodeLib.Data.SqlTools.Sql
{
    public class JoinStatement : FromStatement, IJoinStatement
    {
        public JoinStatement(ISource source, ICondition condition, JoinOperator joinOperator = JoinOperator.Inner)
            : base(source)
        {
            JoinOperator = joinOperator;
            Condition = condition;
        }

        public JoinStatement(string source, ICondition condition, JoinOperator joinOperator = JoinOperator.Inner)
            : base(source)
        {
            JoinOperator = joinOperator;
            Condition = condition;
        }

        public JoinOperator JoinOperator { get; private set; }

        public ICondition Condition { get; private set; }
    }
}
