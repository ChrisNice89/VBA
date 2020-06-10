using AccessCodeLib.Data.Common.Sql;

namespace AccessCodeLib.Data.SqlTools.Sql
{
    public class SelectStatement : FieldsStatement, ISelectStatement
    {
        public SelectStatement()
        {
        }

        public SelectStatement(params IField[] fields) : base(fields)
        {
        }

        public SelectStatement(params string[] fieldNames) : base(fieldNames)
        {
        }

        public override string Key
        {
            get { return StatementKeys.Select; }
        }
    }
}
