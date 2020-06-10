using AccessCodeLib.Data.Common.Sql;

namespace AccessCodeLib.Data.SqlTools.Sql
{
    public class FieldCondition : IFieldCondition
    {
        public FieldCondition()
        {
        }

        public FieldCondition(IField field, RelationalOperators relationalOperator, object value)
        {
            Field = field;
            Operator = relationalOperator;
            Value = value;
        }

        public IField Field { get; set; }

        public RelationalOperators Operator { get; set; }

        public object Value { get; set; }
    }
}