using AccessCodeLib.Data.Common.Sql;

namespace AccessCodeLib.Data.SqlTools.Sql
{
    public class GroupByStatement : FieldsStatement, IGroupByStatement
    {
        public GroupByStatement()
        {
        }

        public GroupByStatement(params IField[] fields)
            : base(fields)
        {
        }

        public GroupByStatement(params string[] fieldNames)
            : base(fieldNames)
        {
        }

        public override string Key
        {
            get { return StatementKeys.GroupBy; }
        }

    }
}