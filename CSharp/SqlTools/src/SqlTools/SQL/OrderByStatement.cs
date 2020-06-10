using AccessCodeLib.Data.Common.Sql;

namespace AccessCodeLib.Data.SqlTools.Sql
{
    public class OrderByStatement : FieldsStatement, IOrderByStatement
    {
        public OrderByStatement()
        {
        }

        public OrderByStatement(params IField[] fields)
            : base(fields)
        {
        }

        public OrderByStatement(params string[] fieldNames)
            : base(fieldNames)
        {
        }

        public override string Key
        {
            get { return StatementKeys.OrderBy; }
        }

    }
}